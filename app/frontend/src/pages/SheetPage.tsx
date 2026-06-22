import { useEffect, useRef, useState, useCallback } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import { io, Socket } from 'socket.io-client';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import { ArrowLeft, Download, Upload, Users, Save, History, X, Info } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';
import { useFonts } from '../fonts/useFonts';

interface OnlineUser { id: string; username: string }

interface ImportState {
  active: boolean;
  progress: number;
  error: string | null;
}

interface CellChange {
  key: string;      // "r_c"
  col: string;      // "A1"
  newVal: string;
  oldVal: string;
}

interface ChangeEntry {
  username: string;
  sheet_index: number;
  summary: string | null;
  changed_cells: CellChange[] | null;
  saved_at: string;
}

// Map of "sheetIndex_r_c" → latest change info for cell highlighting
type CellHighlightMap = Record<string, { username: string; saved_at: string; newVal: string }>;

function buildSheets(sheetMeta: any) {
  return (sheetMeta.sheets || []).map((s: any, i: number) => {
    const data = s.data || {};
    const sheet: any = {
      name: data.name || `Sheet${i + 1}`,
      index: i,
      status: 1,
      celldata: flattenCells(data.cells || {}),
      config: {
        columnlen: numericKeys(data.columnWidths || {}),
        rowlen: numericKeys(data.rowHeights || {}),
        merge: data.merges || {},
        borderInfo: data.borderInfo || [],
        colhidden: data.colhidden || {},
        rowhidden: data.rowhidden || {},
      },
    };

    const filterSelect = data.filterSelect || excelRefToFilterSelect(data.filter?.ref);
    if (filterSelect) {
      sheet.filter_select = filterSelect;
      if (data.filterCriteria) sheet.filter = data.filterCriteria;
    }

    if (data.frozen) sheet.frozen = data.frozen;
    return sheet;
  });
}

function colToNum(letters: string) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function excelRefToFilterSelect(ref?: string) {
  if (!ref) return null;
  const m = /^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/.exec(ref.trim());
  if (!m) return null;
  return {
    row: [Number(m[2]) - 1, Number(m[4]) - 1],
    column: [colToNum(m[1]), colToNum(m[3])],
  };
}

function formatTime(iso: string) {
  const d = new Date(iso);
  return d.toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' });
}

export default function SheetPage() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const { token, isEditor } = useAuth();
  const qc = useQueryClient();
  const socketRef = useRef<Socket | null>(null);
  const [onlineUsers, setOnlineUsers] = useState<OnlineUser[]>([]);
  const [sheets, setSheets] = useState<any[]>([]);
  const [importState, setImportState] = useState<ImportState>({ active: false, progress: 0, error: null });
  const [workbookKey, setWorkbookKey] = useState(0);
  const [saveState, setSaveState] = useState<'idle' | 'saving' | 'saved'>('idle');
  const [isDirty, setIsDirty] = useState(false);
  const [showHistory, setShowHistory] = useState(false);
  const [changelog, setChangelog] = useState<ChangeEntry[]>([]);
  const [cellHighlights, setCellHighlights] = useState<CellHighlightMap>({});
  const [userColors, setUserColors] = useState<Record<string, string>>({});
  // Overlay: absolute rect on the hovered cell + who edited it
  const [hoverOverlay, setHoverOverlay] = useState<{
    left: number; top: number; width: number; height: number;
    color: string; username: string | null;
  } | null>(null);
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  // When navigating from history — show a highlighted border on that cell for 2s
  const [navHighlight, setNavHighlight] = useState<{ r: number; c: number; sheetIdx: number } | null>(null);
  const workbookWrapperRef = useRef<HTMLDivElement>(null);
  const sheetScrollRef = useRef({ top: 0, left: 0 });
  // Canvas top-left relative to wrapper (measured on first mousemove, reset on workbook remount)
  const gridOriginRef = useRef({ top: -1, left: -1 });
  const sheetMetaRef = useRef<any>(null);
  const workbookRef = useRef<any>(null);
  const latestSheetsRef = useRef<any[] | null>(null);
  const lastSavedSheetsRef = useRef<any[] | null>(null);
  // FortuneSheet fires onChange during init — ignore until ready, then treat first call as baseline
  const acceptChangesRef = useRef(false);
  const baselineTakenRef = useRef(false);
  const { version: fontsVersion } = useFonts();

  const editor = isEditor();

  // Load user colors on mount
  useEffect(() => {
    api.get('/auth/user-colors').then((r) => {
      const map: Record<string, string> = {};
      for (const u of r.data) map[u.username] = u.color;
      setUserColors(map);
    }).catch(() => {});
  }, []);

  // Track FortuneSheet scroll in capture phase so we can accurately position the cell overlay
  useEffect(() => {
    const wrapper = workbookWrapperRef.current;
    if (!wrapper) return;
    // Reset grid origin so it's re-measured after workbook re-renders
    gridOriginRef.current = { top: -1, left: -1 };
    sheetScrollRef.current = { top: 0, left: 0 };

    const onScroll = (e: Event) => {
      const el = e.target as HTMLElement;
      if (!el || el === wrapper) return;
      const canScrollV = el.scrollHeight > el.clientHeight;
      const canScrollH = el.scrollWidth > el.clientWidth;
      if (canScrollV) sheetScrollRef.current = { ...sheetScrollRef.current, top: el.scrollTop };
      if (canScrollH) sheetScrollRef.current = { ...sheetScrollRef.current, left: el.scrollLeft };
    };
    wrapper.addEventListener('scroll', onScroll, true);
    return () => wrapper.removeEventListener('scroll', onScroll, true);
  }, [workbookKey]); // re-attach when workbook remounts

  // Reset grid origin measurement when browser zoom changes (devicePixelRatio changes)
  useEffect(() => {
    let lastDpr = window.devicePixelRatio;
    const check = () => {
      if (window.devicePixelRatio !== lastDpr) {
        lastDpr = window.devicePixelRatio;
        gridOriginRef.current = { top: -1, left: -1 };
        setHoverOverlay(null);
      }
    };
    // matchMedia trick: fires when the current dpr-specific query stops matching
    const mq = window.matchMedia(`(resolution: ${lastDpr}dppx)`);
    mq.addEventListener('change', check);
    return () => mq.removeEventListener('change', check);
  }, []);

  useEffect(() => {
    if (fontsVersion > 0) {
      acceptChangesRef.current = false;
      baselineTakenRef.current = false;
      setWorkbookKey((k) => k + 1);
      setTimeout(() => { acceptChangesRef.current = true; }, 800);
    }
  }, [fontsVersion]);

  const { data: sheetMeta, isError: sheetError } = useQuery({
    queryKey: ['sheet', id],
    queryFn: () => api.get(`/spreadsheets/${id}`).then((r) => r.data),
    retry: false,
    staleTime: 0,   // always refetch when navigating back to the page
    gcTime: 0,
  });

  useEffect(() => {
    if (!sheetMeta) return;
    sheetMetaRef.current = sheetMeta;
    // Load changelog to build cell highlights, then build sheets
    api.get(`/spreadsheets/${id}/changelog`).then((r) => {
      const entries: ChangeEntry[] = r.data;
      setChangelog(entries);
      const map: CellHighlightMap = {};
      // Process oldest→newest so newest wins
      for (const entry of [...entries].reverse()) {
        if (!entry.changed_cells) continue;
        for (const cc of entry.changed_cells) {
          map[`${entry.sheet_index}_${cc.key}`] = {
            username: entry.username,
            saved_at: entry.saved_at,
            newVal: cc.newVal,
          };
        }
      }
      setCellHighlights(map);
      const built = buildSheets(sheetMeta);
      const data = built.length ? built : [{ name: 'Sheet1', index: 0, status: 1, celldata: [], config: {} }];
      setSheets(data);
      lastSavedSheetsRef.current = data;
    }).catch(() => {
      const built = buildSheets(sheetMeta);
      const data = built.length ? built : [{ name: 'Sheet1', index: 0, status: 1, celldata: [], config: {} }];
      setSheets(data);
      lastSavedSheetsRef.current = data;
    });
    acceptChangesRef.current = false;
    baselineTakenRef.current = false;
    setIsDirty(false);
    setTimeout(() => { acceptChangesRef.current = true; }, 800);
  }, [sheetMeta, id]);

  // Reload changelog when panel opens (to get latest)
  useEffect(() => {
    if (!showHistory || !id) return;
    api.get(`/spreadsheets/${id}/changelog`).then((r) => setChangelog(r.data));
  }, [showHistory, id]);

  useEffect(() => {
    if (!token) return;
    const socket = io('/', { auth: { token } });
    socketRef.current = socket;
    socket.emit('join-sheet', id);
    socket.on('room-users', setOnlineUsers);
    socket.on('cell-change', ({ changes }) => {
      workbookRef.current?.applyOp?.(changes);
    });
    socket.on('user-color-changed', ({ username, color }: { username: string; color: string }) => {
      setUserColors((prev) => ({ ...prev, [username]: color }));
    });

    socket.on('changelog-update', (entry: ChangeEntry) => {
      setChangelog((prev) => [entry, ...prev].slice(0, 100));
      // Update cell highlights for cells changed by others
      if (entry.changed_cells) {
        setCellHighlights((prev) => {
          const next = { ...prev };
          for (const cc of entry.changed_cells!) {
            next[`${entry.sheet_index}_${cc.key}`] = {
              username: entry.username,
              saved_at: entry.saved_at,
              newVal: cc.newVal,
            };
          }
          return next;
        });
      }
    });
    return () => { socket.disconnect(); };
  }, [id, token]);

  // FortuneSheet header dimensions (row-number column width, column-letter row height)
  const ROW_HEADER_W = 46;
  const COL_HEADER_H = 20;

  // Find the topmost canvas in the workbook — FortuneSheet renders the full grid (including
  // column/row headers) on a single canvas that starts just below the toolbar+formula bar.
  // Cells themselves begin ROW_HEADER_W px to the right and COL_HEADER_H px below the canvas.
  const measureGridOrigin = useCallback(() => {
    const wrapper = workbookWrapperRef.current;
    if (!wrapper) return;
    const wRect = wrapper.getBoundingClientRect();
    const canvases = Array.from(wrapper.querySelectorAll('canvas'));
    if (canvases.length > 0) {
      const topCanvas = canvases.reduce((best, c) =>
        c.getBoundingClientRect().top < best.getBoundingClientRect().top ? c : best
      );
      const cRect = topCanvas.getBoundingClientRect();
      gridOriginRef.current = {
        top:  cRect.top  - wRect.top,
        left: cRect.left - wRect.left,
      };
    } else {
      // No canvas yet — use structural estimates (toolbar≈40px, formula bar≈28px)
      gridOriginRef.current = { top: editor ? 68 : 28, left: 0 };
    }
  }, [editor]);

  // Compute cell rect (relative to workbookWrapper) for a given sheet/row/col
  const getCellRect = useCallback((sheetIdx: number, row: number, col: number) => {
    if (gridOriginRef.current.top < 0) measureGridOrigin();
    if (gridOriginRef.current.top < 0) return null;
    const curSheet = (latestSheetsRef.current ?? sheets)[sheetIdx];
    const colLens = curSheet?.config?.columnlen ?? {};
    const rowLens = curSheet?.config?.rowlen ?? {};
    const DW = 73, DH = 19;
    let xAcc = 0; for (let i = 0; i < col; i++) xAcc += (colLens[i] ?? DW);
    let yAcc = 0; for (let i = 0; i < row; i++) yAcc += (rowLens[i] ?? DH);
    const { top: oT, left: oL } = gridOriginRef.current;
    const { top: st, left: sl } = sheetScrollRef.current;
    return {
      left:   oL + ROW_HEADER_W - sl + xAcc,
      top:    oT + COL_HEADER_H - st + yAcc,
      width:  colLens[col] ?? DW,
      height: rowLens[row] ?? DH,
    };
  }, [sheets, ROW_HEADER_W, COL_HEADER_H, measureGridOrigin]);

  const handleWorkbookMouseMove = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    const wrapper = workbookWrapperRef.current;
    if (!wrapper) return;
    if (gridOriginRef.current.top < 0) measureGridOrigin();

    const wRect = wrapper.getBoundingClientRect();
    const x = e.clientX - wRect.left;
    const y = e.clientY - wRect.top;

    const { top: oT, left: oL } = gridOriginRef.current;
    const { top: st, left: sl } = sheetScrollRef.current;

    // Offset inside the cell grid (excluding row/column headers)
    const adjX = x - oL - ROW_HEADER_W + sl;
    const adjY = y - oT - COL_HEADER_H + st;

    if (adjX < 0 || adjY < 0) { setHoverOverlay(null); return; }

    const curSheet = (latestSheetsRef.current ?? sheets)[activeSheetIdx];
    const colLens = curSheet?.config?.columnlen ?? {};
    const rowLens = curSheet?.config?.rowlen ?? {};
    const DW = 73, DH = 19;

    let col = 0, xAcc = 0;
    while (col < 500) { const w = colLens[col] ?? DW; if (xAcc + w > adjX) break; xAcc += w; col++; }
    let row = 0, yAcc = 0;
    while (row < 2000) { const h = rowLens[row] ?? DH; if (yAcc + h > adjY) break; yAcc += h; row++; }

    const hKey = `${activeSheetIdx}_${row}_${col}`;
    const change = cellHighlights[hKey];
    const color = change ? (userColors[change.username] ?? '#3B82F6') : '#94a3b8';

    setHoverOverlay({
      left:   oL + ROW_HEADER_W - sl + xAcc,
      top:    oT + COL_HEADER_H - st + yAcc,
      width:  colLens[col] ?? DW,
      height: rowLens[row] ?? DH,
      color, username: change?.username ?? null,
    });
  }, [sheets, activeSheetIdx, cellHighlights, userColors, measureGridOrigin, ROW_HEADER_W, COL_HEADER_H]);

  const handleWorkbookMouseLeave = useCallback(() => {
    setHoverOverlay(null);
  }, []);

  // Navigate to cell from history click — scroll there and highlight with a pulsing overlay
  const navigateToCell = useCallback((sheetIndex: number, r: number, c: number) => {
    setActiveSheetIdx(sheetIndex);
    setNavHighlight({ r, c, sheetIdx: sheetIndex });
    setTimeout(() => setNavHighlight(null), 2500);

    const wrapper = workbookWrapperRef.current;
    if (!wrapper) return;
    const curSheet = (latestSheetsRef.current ?? sheets)[sheetIndex];
    const colLens = curSheet?.config?.columnlen ?? {};
    const rowLens = curSheet?.config?.rowlen ?? {};
    const DW = 73, DH = 19;
    let scrollX = 0; for (let i = 0; i < c; i++) scrollX += (colLens[i] ?? DW);
    let scrollY = 0; for (let i = 0; i < r; i++) scrollY += (rowLens[i] ?? DH);

    // Scroll all scrollable containers within the workbook
    const all = Array.from(wrapper.querySelectorAll('*')) as HTMLElement[];
    for (const el of all) {
      if (el.scrollHeight > el.clientHeight + 2) el.scrollTop = Math.max(0, scrollY - 80);
      if (el.scrollWidth  > el.clientWidth  + 2) el.scrollLeft = Math.max(0, scrollX - 80);
    }
  }, [sheets]);

  const computeChangedCells = useCallback((current: any[], prev: any[] | null): CellChange[] => {
    if (!prev) return [];
    const changes: CellChange[] = [];
    for (const sheet of current) {
      const si = sheet.index ?? 0;
      const prevSheet = prev.find((s: any) => (s.index ?? 0) === si);
      const currCells = cellsFromSheet(sheet);
      const prevCells = prevSheet ? cellsFromSheet(prevSheet) : {};
      const allKeys = new Set([...Object.keys(currCells), ...Object.keys(prevCells)]);
      for (const key of allKeys) {
        const cv = String(currCells[key]?.v ?? currCells[key]?.m ?? '').trim();
        const pv = String(prevCells[key]?.v ?? prevCells[key]?.m ?? '').trim();
        // Skip if value unchanged or both empty
        if (cv === pv || (cv === '' && pv === '')) continue;
        const [r, c] = key.split('_').map(Number);
        changes.push({ key, col: `${colName(c)}${r + 1}`, newVal: cv.slice(0, 100), oldVal: pv.slice(0, 100) });
      }
    }
    return changes.slice(0, 200);
  }, []);

  const buildSummary = useCallback((current: any[]): string => {
    const prev = lastSavedSheetsRef.current;
    if (!prev) return '';
    let changedCount = 0;
    const examples: string[] = [];
    for (const sheet of current) {
      const prevSheet = prev.find((s: any) => (s.index ?? s.name) === (sheet.index ?? sheet.name));
      const currCells = cellsFromSheet(sheet);
      const prevCells = prevSheet ? cellsFromSheet(prevSheet) : {};
      const allKeys = new Set([...Object.keys(currCells), ...Object.keys(prevCells)]);
      for (const key of allKeys) {
        const cv = currCells[key]?.v ?? currCells[key]?.m ?? '';
        const pv = prevCells[key]?.v ?? prevCells[key]?.m ?? '';
        if (String(cv) !== String(pv)) {
          changedCount++;
          if (examples.length < 3) {
            const [r, c] = key.split('_').map(Number);
            const col = colName(c);
            const val = String(cv).slice(0, 30);
            examples.push(`${col}${r + 1}${val ? `="${val}"` : ' (удалено)'}`);
          }
        }
      }
    }
    if (!changedCount) return '';
    return `Изменено ячеек: ${changedCount}${examples.length ? ` (${examples.join(', ')})` : ''}`;
  }, []);

  const saveAll = useCallback((allSheets: any[], summary?: string, changedCells?: CellChange[]) => {
    (allSheets || []).forEach((s, i) => {
      const data = {
        name: s.name,
        cells: cellsFromSheet(s),
        columnWidths: s.config?.columnlen || {},
        rowHeights: s.config?.rowlen || {},
        merges: s.config?.merge || {},
        borderInfo: s.config?.borderInfo || [],
        filterSelect: s.filter_select || null,
        filterCriteria: s.filter || null,
        frozen: s.frozen || null,
        colhidden: s.config?.colhidden || {},
        rowhidden: s.config?.rowhidden || {},
      };
      socketRef.current?.emit('save-sheet', {
        sheetId: id,
        sheetIndex: s.index ?? i,
        data,
        summary: summary || null,
        changedCells: changedCells || null,
      });
    });
  }, [id]);

  const handleSaveNow = useCallback(() => {
    const all = latestSheetsRef.current || sheets;
    const summary = buildSummary(all);
    const changedCells = computeChangedCells(all, lastSavedSheetsRef.current);
    setSaveState('saving');
    saveAll(all, summary, changedCells);
    lastSavedSheetsRef.current = all;
    setIsDirty(false);
    setSaveState('saved');
    setTimeout(() => setSaveState('idle'), 2000);
  }, [saveAll, buildSummary, sheets]);

  const handleChange = useCallback((allSheets: any) => {
    if (!editor || !allSheets?.length) return;
    if (!acceptChangesRef.current) return;
    // First onChange after init timeout: treat as baseline (FortuneSheet may still be normalizing)
    if (!baselineTakenRef.current) {
      baselineTakenRef.current = true;
      latestSheetsRef.current = allSheets;
      lastSavedSheetsRef.current = allSheets;
      return;
    }
    latestSheetsRef.current = allSheets;
    setIsDirty(true);
    // Broadcast to other users for real-time collaboration (no auto-save)
    socketRef.current?.emit('cell-change', { sheetId: id, changes: allSheets });
  }, [id, editor]);

  const handleExport = async () => {
    const res = await api.get(`/excel/${id}/export`, { responseType: 'blob' });
    const url = URL.createObjectURL(res.data);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${sheetMeta?.name || 'table'}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';

    const jobId = `job_${Date.now()}`;
    setImportState({ active: true, progress: 2, error: null });

    const evtSource = new EventSource(
      `/api/excel/${id}/import-progress?jobId=${jobId}&token=${encodeURIComponent(token || '')}`
    );

    evtSource.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        setImportState((s) => ({ ...s, progress: data.progress }));
        if (data.done) {
          evtSource.close();
          if (data.error) {
            setImportState({ active: false, progress: 0, error: data.error });
          } else {
            api.get(`/spreadsheets/${id}`).then((r) => {
              const fresh = buildSheets(r.data);
              setSheets(fresh.length ? fresh : [{ name: 'Sheet1', index: 0, status: 1, celldata: [], config: {} }]);
              setWorkbookKey((k) => k + 1);
              setImportState({ active: false, progress: 0, error: null });
              setIsDirty(false);
            });
          }
        }
      } catch { /* ignore parse errors */ }
    };
    evtSource.onerror = () => evtSource.close();

    try {
      const form = new FormData();
      form.append('file', file);
      await api.post(`/excel/${id}/import?jobId=${jobId}`, form);
    } catch (err: any) {
      evtSource.close();
      setImportState({ active: false, progress: 0, error: err.response?.data?.error || 'Ошибка импорта' });
    }
  };

  if (sheetError) {
    return (
      <div className="flex flex-col items-center justify-center h-screen gap-4 text-gray-500">
        <p>Нет доступа к этой таблице.</p>
        <button onClick={() => navigate('/')} className="text-blue-600 hover:underline">На главную</button>
      </div>
    );
  }

  if (!sheets.length) {
    return <div className="flex items-center justify-center h-screen text-gray-400">Загрузка...</div>;
  }

  return (
    <div className="flex flex-col h-screen bg-white">
      {/* Import progress overlay */}
      {importState.active && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center">
          <div className="bg-white rounded-2xl shadow-xl p-8 w-80">
            <h3 className="font-semibold text-gray-800 mb-4 text-center">Импорт файла Excel</h3>
            <div className="w-full bg-gray-100 rounded-full h-3 mb-3 overflow-hidden">
              <div
                className="bg-blue-500 h-3 rounded-full transition-all duration-200"
                style={{ width: `${importState.progress}%` }}
              />
            </div>
            <p className="text-center text-sm text-gray-500">
              {importState.progress < 95
                ? `${importState.progress}% — обработка ячеек...`
                : 'Сохранение в базу данных...'}
            </p>
          </div>
        </div>
      )}

      {importState.error && (
        <div className="fixed top-4 right-4 z-50 bg-red-50 border border-red-200 rounded-xl px-4 py-3 shadow-lg max-w-sm">
          <p className="text-red-600 text-sm font-medium">Ошибка импорта</p>
          <p className="text-red-500 text-xs mt-1">{importState.error}</p>
          <button onClick={() => setImportState((s) => ({ ...s, error: null }))}
            className="text-xs text-red-400 underline mt-2 block">Закрыть</button>
        </div>
      )}

      {/* Header */}
      <header className="bg-white border-b px-3 py-2 flex items-center gap-2 shrink-0 flex-wrap sm:flex-nowrap">
        <button onClick={() => navigate('/')} className="p-1 rounded hover:bg-gray-100 shrink-0">
          <ArrowLeft size={18} />
        </button>
        <h2 className="font-semibold text-gray-800 flex-1 min-w-0 truncate text-sm sm:text-base">
          {sheetMeta?.name}
        </h2>

        <div className="flex items-center gap-1.5 shrink-0 flex-wrap justify-end">
          {onlineUsers.length > 0 && (
            <div className="hidden sm:flex items-center gap-1 text-xs text-gray-500 max-w-[160px] overflow-hidden">
              <Users size={13} className="shrink-0" />
              {onlineUsers.slice(0, 3).map((u) => (
                <span key={u.id} className="bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded truncate">{u.username}</span>
              ))}
              {onlineUsers.length > 3 && <span className="text-gray-400 text-xs">+{onlineUsers.length - 3}</span>}
            </div>
          )}

          {editor ? (
            <>
              <button
                onClick={handleSaveNow}
                title="Сохранить все изменения (обязательно нажмите после редактирования)"
                className={`flex items-center gap-1 text-xs sm:text-sm rounded-lg px-2 sm:px-3 py-1.5 border transition font-medium ${
                  saveState === 'saved'
                    ? 'border-green-300 bg-green-50 text-green-600'
                    : isDirty
                    ? 'border-orange-300 bg-orange-50 text-orange-600 animate-pulse'
                    : 'border-gray-200 text-gray-600 hover:text-gray-800 hover:bg-gray-50'
                }`}
              >
                <Save size={14} />
                <span className="hidden xs:inline">
                  {saveState === 'saved' ? 'Сохранено' : 'Сохранить'}
                </span>
              </button>
              <label
                title="Импорт .xlsx"
                className="flex items-center gap-1 cursor-pointer text-xs sm:text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-2 sm:px-3 py-1.5 hover:bg-gray-50 transition"
              >
                <Upload size={14} />
                <span className="hidden sm:inline">Импорт</span>
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport} />
              </label>
              <button
                onClick={handleExport}
                title="Экспорт .xlsx"
                className="flex items-center gap-1 text-xs sm:text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-2 sm:px-3 py-1.5 hover:bg-gray-50 transition"
              >
                <Download size={14} />
                <span className="hidden sm:inline">Экспорт</span>
              </button>
            </>
          ) : (
            <button
              onClick={handleExport}
              title="Экспорт .xlsx"
              className="flex items-center gap-1 text-xs sm:text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-2 sm:px-3 py-1.5 hover:bg-gray-50 transition"
            >
              <Download size={14} />
              <span className="hidden sm:inline">Экспорт</span>
            </button>
          )}

          <button
            onClick={() => setShowHistory((v) => !v)}
            title="История изменений"
            className={`flex items-center gap-1 text-xs sm:text-sm rounded-lg px-2 sm:px-3 py-1.5 border transition ${
              showHistory ? 'border-blue-300 bg-blue-50 text-blue-600' : 'border-gray-200 text-gray-600 hover:bg-gray-50'
            }`}
          >
            <History size={14} />
            <span className="hidden sm:inline">История</span>
          </button>
        </div>
      </header>

      {/* Reader banner */}
      {!editor && (
        <div className="bg-blue-50 border-b border-blue-100 px-4 py-2 flex items-center gap-2 text-sm text-blue-700 shrink-0">
          <Info size={15} className="shrink-0" />
          <span>Вы в режиме чтения — редактирование недоступно.</span>
        </div>
      )}

      {/* Dirty hint for editors */}
      {editor && isDirty && saveState === 'idle' && (
        <div className="bg-orange-50 border-b border-orange-100 px-4 py-1.5 flex items-center gap-2 text-xs text-orange-700 shrink-0">
          <Info size={14} className="shrink-0" />
          <span>Есть несохранённые изменения — нажмите кнопку <strong>«Сохранить»</strong> чтобы не потерять их.</span>
        </div>
      )}

      {/* Cell hover overlay — absolutely positioned over the exact cell */}

      {/* Main content area */}
      <div className="flex flex-1 overflow-hidden">
        <div
          className="flex-1 overflow-hidden relative"
          ref={workbookWrapperRef}
          onMouseEnter={() => { gridOriginRef.current = { top: -1, left: -1 }; }}
          onMouseMove={handleWorkbookMouseMove}
          onMouseLeave={handleWorkbookMouseLeave}
        >
          {/* Hover overlay — highlights the cell under the cursor */}
          {hoverOverlay && (
            <div
              style={{
                position: 'absolute',
                left: hoverOverlay.left,
                top: hoverOverlay.top,
                width: hoverOverlay.width,
                height: hoverOverlay.height,
                background: hoverOverlay.color + '33',
                border: `1px solid ${hoverOverlay.color}88`,
                pointerEvents: 'none',
                zIndex: 50,
                display: 'flex',
                alignItems: 'flex-end',
                overflow: 'hidden',
                boxSizing: 'border-box',
              }}
            >
              {hoverOverlay.username && (
                <span style={{
                  fontSize: 9,
                  lineHeight: '10px',
                  color: hoverOverlay.color,
                  padding: '0 2px 1px',
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  maxWidth: '100%',
                  fontWeight: 600,
                }}>
                  {hoverOverlay.username}
                </span>
              )}
            </div>
          )}

          {/* Navigation highlight — shown 2.5s after clicking a history entry */}
          {navHighlight && (() => {
            const rect = getCellRect(navHighlight.sheetIdx, navHighlight.r, navHighlight.c);
            if (!rect) return null;
            const hKey = `${navHighlight.sheetIdx}_${navHighlight.r}_${navHighlight.c}`;
            const change = cellHighlights[hKey];
            const color = change ? (userColors[change.username] ?? '#3B82F6') : '#3B82F6';
            return (
              <div
                style={{
                  position: 'absolute',
                  left: rect.left,
                  top: rect.top,
                  width: rect.width,
                  height: rect.height,
                  background: color + '44',
                  border: `2px solid ${color}`,
                  pointerEvents: 'none',
                  zIndex: 51,
                  boxSizing: 'border-box',
                  animation: 'twSheetPulse 0.6s ease-in-out 3',
                }}
              />
            );
          })()}

          <Workbook
            key={workbookKey}
            ref={workbookRef}
            data={sheets}
            lang="ru"
            onChange={handleChange}
            showToolbar={editor}
            showFormulaBar
            allowEdit={editor}
            toolbarItems={[
              'undo','redo','|',
              'format-painter','clear-format','|',
              'font','|',
              'font-size','|',
              'bold','italic','strike-through','underline','|',
              'font-color','background','|',
              'border','merge-cell','|',
              'horizontal-align','vertical-align','text-wrap','text-rotation','|',
              'currency-format','percentage-format','number-decrease','number-increase','format','|',
              'freeze','filter','conditionFormat','|',
              'quick-formula','|',
              'comment','link','image','|',
              'search','screenshot',
            ]}
          />
        </div>

        {/* History sidebar — overlay on mobile, sidebar on desktop */}
        {showHistory && (
          <>
            {/* Mobile backdrop */}
            <div
              className="fixed inset-0 bg-black/30 z-20 sm:hidden"
              onClick={() => setShowHistory(false)}
            />
            <div className="fixed right-0 top-0 bottom-0 w-80 max-w-[90vw] z-30 sm:static sm:z-auto sm:w-72 sm:max-w-none border-l bg-white flex flex-col shrink-0 overflow-hidden shadow-xl sm:shadow-none">
              <div className="flex items-center justify-between px-3 py-3 border-b shrink-0">
                <span className="text-sm font-semibold text-gray-700 flex items-center gap-1.5">
                  <History size={14} /> История изменений
                </span>
                <button onClick={() => setShowHistory(false)} className="p-1.5 rounded hover:bg-gray-100 text-gray-400">
                  <X size={16} />
                </button>
              </div>
              <div className="flex-1 overflow-y-auto">
                {changelog.length === 0 ? (
                  <p className="text-xs text-gray-400 text-center py-8 px-3">Изменений пока нет</p>
                ) : (
                  <ul className="divide-y">
                    {changelog.map((e, i) => {
                      const uColor = userColors[e.username] ?? '#3B82F6';
                      return (
                        <li key={i} className="px-3 py-3">
                          <div className="flex items-center gap-1.5 mb-0.5">
                            <span
                              className="inline-flex items-center justify-center w-6 h-6 rounded-full text-white text-xs font-bold shrink-0"
                              style={{ background: uColor }}
                            >
                              {e.username[0].toUpperCase()}
                            </span>
                            <span className="text-sm font-medium text-gray-800 truncate">{e.username}</span>
                          </div>
                          {e.summary && (
                            <p className="text-xs text-gray-600 mt-1 ml-7 break-words">{e.summary}</p>
                          )}
                          {e.changed_cells && e.changed_cells.length > 0 && (
                            <div className="ml-7 mt-1.5 space-y-0.5">
                              {e.changed_cells.slice(0, 8).map((cc, ci) => {
                                const [r, c] = cc.key.split('_').map(Number);
                                return (
                                  <button
                                    key={ci}
                                    className="flex items-start gap-1 text-xs w-full text-left hover:bg-gray-50 rounded px-1 -mx-1 py-0.5 transition group"
                                    onClick={() => {
                                      navigateToCell(e.sheet_index, r, c);
                                    }}
                                    title={`Перейти к ячейке ${cc.col}`}
                                  >
                                    <span className="font-mono font-semibold shrink-0 group-hover:underline" style={{ color: uColor }}>{cc.col}</span>
                                    <span className="text-gray-400 shrink-0">→</span>
                                    <span className="text-gray-700 break-all">
                                      {cc.newVal ? `"${cc.newVal.slice(0, 40)}"` : <span className="italic text-gray-400">удалено</span>}
                                    </span>
                                  </button>
                                );
                              })}
                              {e.changed_cells.length > 8 && (
                                <p className="text-xs text-gray-400">ещё {e.changed_cells.length - 8} ячеек…</p>
                              )}
                            </div>
                          )}
                          <p className="text-xs text-gray-400 mt-0.5 ml-7">{formatTime(e.saved_at)}</p>
                        </li>
                      );
                    })}
                  </ul>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function colName(c: number): string {
  let name = '';
  c += 1;
  while (c > 0) {
    name = String.fromCharCode(65 + ((c - 1) % 26)) + name;
    c = Math.floor((c - 1) / 26);
  }
  return name;
}

function numericKeys(obj: Record<string, any>) {
  const result: Record<number, any> = {};
  Object.entries(obj).forEach(([k, v]) => { result[Number(k)] = v; });
  return result;
}

function flattenCells(cells: Record<string, any>) {
  return Object.entries(cells).map(([key, v]) => {
    const [r, c] = key.split('_').map(Number);
    return { r, c, v };
  });
}

function unflattenCells(celldata: any[]) {
  const cells: Record<string, any> = {};
  (celldata || []).forEach(({ r, c, v }) => {
    if (v !== undefined && v !== null) cells[`${r}_${c}`] = v;
  });
  return cells;
}

function cellsFromSheet(data: any): Record<string, any> {
  const matrix = data?.data;
  if (Array.isArray(matrix) && matrix.length) {
    const cells: Record<string, any> = {};
    for (let r = 0; r < matrix.length; r++) {
      const row = matrix[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        if (row[c] != null) cells[`${r}_${c}`] = row[c];
      }
    }
    return cells;
  }
  return unflattenCells(data?.celldata || []);
}
