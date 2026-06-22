import { useEffect, useRef, useState, useCallback, useMemo } from 'react';
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
      // Stable id (shared across clients) so live-presence cursors map to the same sheet —
      // otherwise FortuneSheet auto-generates a different id per client.
      id: `sheet_${i}`,
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

// Stable, readable color derived from a username — used when a user has no color
// in the DB (or hasn't loaded yet). Deterministic so the same user always looks the
// same, instead of everyone falling back to one shared blue.
function colorFromUsername(name: string): string {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) | 0;
  const hue = Math.abs(h) % 360;
  return `hsl(${hue}, 65%, 45%)`;
}

// Resolve the display color for an edit author: DB color if present, else a stable
// per-username fallback (never the old shared #3B82F6).
function userColor(colors: Record<string, string>, name: string): string {
  return colors[name] || colorFromUsername(name);
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
  // Bumped whenever the cell-rect map is rebuilt (scroll/zoom/redraw) → re-renders persistent overlays
  const [mapVersion, setMapVersion] = useState(0);
  // When navigating from history — pulse this resolved absolute rect for 2.5s
  const [navHighlight, setNavHighlight] = useState<{
    left: number; top: number; width: number; height: number; color: string;
  } | null>(null);
  const workbookWrapperRef = useRef<HTMLDivElement>(null);
  // Exact CSS-pixel rects of every VISIBLE cell, captured from FortuneSheet's afterRenderCell hook.
  // Coords are relative to the .fortune-sheet-canvas top-left (already CSS px — FortuneSheet's
  // canvas context is pre-scaled by devicePixelRatio, so no ratio conversion is needed).
  // The map is rebuilt on every redraw (scroll/zoom/resize), so positions are always current.
  const cellRectMapRef = useRef<Map<string, { x: number; y: number; w: number; h: number }>>(new Map());
  const pendingRectMapRef = useRef<Map<string, { x: number; y: number; w: number; h: number }>>(new Map());
  const renderBatchTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  // Mirror of activeSheetIdx for use inside stable useMemo hooks
  const activeSheetIdxRef = useRef(0);
  // Keep ref in sync so stable hooks can read current sheet index
  useEffect(() => { activeSheetIdxRef.current = activeSheetIdx; }, [activeSheetIdx]);
  // True when the active sheet has any edited cells → gates per-frame overlay re-renders
  const hasOverlaysRef = useRef(false);
  // Last-known rect per edited cell, for smooth fade-out when it scrolls out of view
  const overlaySeenRef = useRef<Map<string, { left: number; top: number; w: number; h: number; t: number }>>(new Map());
  // Cell key the cursor is currently over → skip redundant hover state updates
  const lastHoverKeyRef = useRef<string | null>(null);
  useEffect(() => {
    const prefix = `${activeSheetIdx}_`;
    hasOverlaysRef.current = Object.keys(cellHighlights).some((k) => k.startsWith(prefix));
  }, [cellHighlights, activeSheetIdx]);
  // Drop cached overlay positions when switching sheets (avoid ghosts from the previous sheet)
  useEffect(() => { overlaySeenRef.current = new Map(); }, [activeSheetIdx]);
  const sheetMetaRef = useRef<any>(null);
  const workbookRef = useRef<any>(null);
  const latestSheetsRef = useRef<any[] | null>(null);
  const lastSavedSheetsRef = useRef<any[] | null>(null);
  // Always-current mirror of `sheets` so navigateToCell can read column widths / row
  // heights without a stale closure (works for readers too, where latestSheetsRef is null).
  const sheetsRef = useRef<any[]>([]);
  useEffect(() => { sheetsRef.current = sheets; }, [sheets]);
  // Always-current mirror of userColors / online users for use inside stable socket handlers.
  const userColorsRef = useRef<Record<string, string>>({});
  const onlineUsersRef = useRef<OnlineUser[]>([]);
  // FortuneSheet fires onChange during init — ignore until ready, then treat first call as baseline
  const acceptChangesRef = useRef(false);
  const baselineTakenRef = useRef(false);
  const { version: fontsVersion } = useFonts();

  const editor = isEditor();

  useEffect(() => { userColorsRef.current = userColors; }, [userColors]);
  useEffect(() => { onlineUsersRef.current = onlineUsers; }, [onlineUsers]);

  // Local draft (autosaved to localStorage) so unsaved work survives a tab crash/close.
  const draftKey = id ? `tableweb_draft_${id}` : '';
  const [draftInfo, setDraftInfo] = useState<{ savedAt: string } | null>(null);
  const isDirtyRef = useRef(false);
  useEffect(() => { isDirtyRef.current = isDirty; }, [isDirty]);

  // Warn before leaving the tab while there are unsaved changes (manual save only).
  useEffect(() => {
    if (!editor) return;
    const handler = (e: BeforeUnloadEvent) => {
      if (isDirtyRef.current) { e.preventDefault(); e.returnValue = ''; }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [editor]);

  // Autosave a local draft every 5s while there are unsaved edits. Cleared on a real save.
  useEffect(() => {
    if (!editor || !draftKey) return;
    const iv = setInterval(() => {
      if (!isDirtyRef.current) return;
      const all = latestSheetsRef.current;
      if (!all) return;
      try {
        localStorage.setItem(draftKey, JSON.stringify({ savedAt: new Date().toISOString(), sheets: all }));
      } catch { /* quota exceeded — skip this tick */ }
    }, 5000);
    return () => clearInterval(iv);
  }, [editor, draftKey]);

  // On open, surface any leftover draft (its mere presence means it was never saved).
  useEffect(() => {
    if (!editor || !draftKey) return;
    try {
      const raw = localStorage.getItem(draftKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed?.sheets?.length) setDraftInfo({ savedAt: parsed.savedAt });
      }
    } catch { /* ignore corrupt draft */ }
  }, [editor, draftKey]);

  const restoreDraft = useCallback(() => {
    try {
      const raw = localStorage.getItem(draftKey);
      if (!raw) { setDraftInfo(null); return; }
      const parsed = JSON.parse(raw);
      if (parsed?.sheets?.length) {
        acceptChangesRef.current = false;
        baselineTakenRef.current = false;
        setSheets(parsed.sheets);
        latestSheetsRef.current = parsed.sheets;
        setWorkbookKey((k) => k + 1);
        setIsDirty(true);
        setTimeout(() => { acceptChangesRef.current = true; }, 800);
      }
    } catch { /* ignore */ }
    setDraftInfo(null);
  }, [draftKey]);

  const discardDraft = useCallback(() => {
    try { localStorage.removeItem(draftKey); } catch { /* ignore */ }
    setDraftInfo(null);
  }, [draftKey]);

  // Load user colors on mount
  useEffect(() => {
    api.get('/auth/user-colors').then((r) => {
      const map: Record<string, string> = {};
      for (const u of r.data) map[u.username] = u.color;
      setUserColors(map);
    }).catch(() => {});
  }, []);

  // Clear the cell-rect map on workbook remount (it refills from afterRenderCell)
  useEffect(() => {
    cellRectMapRef.current = new Map();
    pendingRectMapRef.current = new Map();
    overlaySeenRef.current = new Map();
    setHoverOverlay(null);
  }, [workbookKey]);

  // Cancel any pending overlay-update frame on unmount
  useEffect(() => () => {
    if (renderBatchTimerRef.current !== null) {
      cancelAnimationFrame(renderBatchTimerRef.current as unknown as number);
    }
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

  // Hide stale hover overlay when the history panel toggles (canvas shifts/redraws)
  useEffect(() => { setHoverOverlay(null); }, [showHistory]);

  useEffect(() => {
    if (!token) return;
    const socket = io('/', { auth: { token } });
    socketRef.current = socket;
    socket.emit('join-sheet', id);
    socket.on('room-users', (users: OnlineUser[]) => {
      // Remove presence cursors of anyone who just left the room
      const goneUsernames = onlineUsersRef.current
        .filter((p) => !users.some((u) => u.username === p.username))
        .map((p) => p.username);
      if (goneUsernames.length) {
        try { workbookRef.current?.removePresences?.(goneUsernames.map((u) => ({ userId: u }))); } catch { /* ignore */ }
      }
      setOnlineUsers(users);
    });
    socket.on('cell-change', ({ changes }) => {
      workbookRef.current?.applyOp?.(changes);
    });
    // Live presence: show other users' selection cursors (FortuneSheet addPresences).
    socket.on('presence', ({ username, tabId, r, c }: { username: string; tabId: string; r: number; c: number }) => {
      const wb = workbookRef.current;
      if (!wb?.addPresences) return;
      const color = userColor(userColorsRef.current, username);
      try {
        wb.removePresences?.([{ userId: username }]);
        wb.addPresences([{ userId: username, username, color, selection: { r, c }, sheetId: tabId }]);
      } catch { /* ignore */ }
    });
    socket.on('user-color-changed', ({ username, color }: { username: string; color: string }) => {
      setUserColors((prev) => ({ ...prev, [username]: color }));
    });

    socket.on('changelog-update', (entry: ChangeEntry) => {
      setChangelog((prev) => {
        // Deduplicate: skip if an entry with same username and timestamp already exists
        const isDup = prev.some(
          (e) => e.username === entry.username && e.saved_at === entry.saved_at
        );
        if (isDup) return prev;
        return [entry, ...prev].slice(0, 100);
      });
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

  // The single FortuneSheet grid canvas. afterRenderCell coords are CSS px relative to its top-left.
  const getCanvas = useCallback((): HTMLCanvasElement | null => {
    return (workbookWrapperRef.current?.querySelector('canvas.fortune-sheet-canvas')
      ?? workbookWrapperRef.current?.querySelector('canvas')) as HTMLCanvasElement | null;
  }, []);

  // Canvas top-left relative to the workbook wrapper, measured fresh (handles panel toggle/scroll).
  const getCanvasOrigin = useCallback((): { left: number; top: number } | null => {
    const wrapper = workbookWrapperRef.current;
    const canvas = getCanvas();
    if (!wrapper || !canvas) return null;
    const wRect = wrapper.getBoundingClientRect();
    const cRect = canvas.getBoundingClientRect();
    return { left: cRect.left - wRect.left, top: cRect.top - wRect.top };
  }, [getCanvas]);

  const handleWorkbookMouseMove = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    const canvas = getCanvas();
    const origin = getCanvasOrigin();
    if (!canvas || !origin) { setHoverOverlay(null); return; }

    const cRect = canvas.getBoundingClientRect();
    // Mouse position relative to the canvas top-left, in CSS px — same space as the map rects
    const cx = e.clientX - cRect.left;
    const cy = e.clientY - cRect.top;

    // Find the visible cell whose exact rect (from FortuneSheet) contains the cursor.
    // The map holds every on-screen cell, so empty cells work too.
    const map = cellRectMapRef.current;
    for (const [key, rect] of map) {
      if (cx >= rect.x && cx < rect.x + rect.w && cy >= rect.y && cy < rect.y + rect.h) {
        // Optimization: cursor still inside the same cell → nothing to update
        if (lastHoverKeyRef.current === key) return;
        lastHoverKeyRef.current = key;
        const [row, col] = key.split('_').map(Number);
        const hKey = `${activeSheetIdx}_${row}_${col}`;
        // Edited cells are already decorated permanently → only neutral-gray hover on others
        if (cellHighlights[hKey]) { setHoverOverlay(null); return; }
        setHoverOverlay({
          left: origin.left + rect.x, top: origin.top + rect.y,
          width: rect.w,             height: rect.h,
          color: '#94a3b8', username: null,
        });
        return;
      }
    }
    if (lastHoverKeyRef.current !== null) { lastHoverKeyRef.current = null; setHoverOverlay(null); }
  }, [activeSheetIdx, cellHighlights, getCanvas, getCanvasOrigin]);

  const handleWorkbookMouseLeave = useCallback(() => {
    lastHoverKeyRef.current = null;
    setHoverOverlay(null);
  }, []);

  // Navigate to cell from history click — center it on screen, then pulse-highlight it.
  const navigateToCell = useCallback((sheetIndex: number, r: number, c: number) => {
    setActiveSheetIdx(sheetIndex);
    setHoverOverlay(null);
    setNavHighlight(null);

    const wb = workbookRef.current as any;
    // Native selection — FortuneSheet draws its own highlight box exactly on the cell.
    try { wb?.setSelection?.([{ row: [r, r], column: [c, c] }]); } catch (_) {}

    const wrap = workbookWrapperRef.current;
    const sbx = wrap?.querySelector('.luckysheet-scrollbar-x') as HTMLElement | null;
    const sby = wrap?.querySelector('.luckysheet-scrollbar-y') as HTMLElement | null;

    // Scroll the cell to the viewport center with a PURE computed scroll: sum the column
    // widths / row heights before the target to get its absolute content offset, place its
    // center at the scrollbar's mid-point. Verified live to land dead-centre even with
    // frozen rows/columns — the scrollbar works in full content px, so the frozen pane needs
    // no special-casing. (The earlier measure-and-nudge refinement is gone: it double-applied
    // a stale rect and overshot, throwing the view far past the target on distant cells.)
    const sheet = sheetsRef.current[sheetIndex] || sheetsRef.current[0];
    const colLen: Record<number, number> = sheet?.config?.columnlen || {};
    const rowLen: Record<number, number> = sheet?.config?.rowlen || {};
    const DEFW = 73, DEFH = 19; // FortuneSheet defaults for cells without an explicit size
    let contentX = 0;
    for (let k = 0; k < c; k++) contentX += colLen[k] ?? DEFW;
    let contentY = 0;
    for (let k = 0; k < r; k++) contentY += rowLen[k] ?? DEFH;
    const cellW = colLen[c] ?? DEFW;
    const cellH = rowLen[r] ?? DEFH;
    if (sbx) sbx.scrollLeft = Math.max(0, contentX + cellW / 2 - sbx.clientWidth / 2);
    if (sby) sby.scrollTop = Math.max(0, contentY + cellH / 2 - sby.clientHeight / 2);

    // After the scroll has rendered, read the cell's exact on-screen rect and pulse-highlight
    // it. Best-effort: a few retries while the redraw settles, then give up quietly.
    let tries = 0;
    const tick = () => {
      tries += 1;
      const rect = cellRectMapRef.current.get(`${r}_${c}`);
      const origin = getCanvasOrigin();
      if (rect && origin) {
        const change = cellHighlights[`${sheetIndex}_${r}_${c}`];
        const color = change ? userColor(userColors, change.username) : '#3B82F6';
        setNavHighlight({
          left: origin.left + rect.x, top: origin.top + rect.y,
          width: rect.w, height: rect.h, color,
        });
        setTimeout(() => setNavHighlight(null), 2500);
      } else if (tries < 25) {
        setTimeout(tick, 70);
      }
    };
    setTimeout(tick, 150);
  }, [cellHighlights, userColors, getCanvasOrigin]);

  // Note: change detection (diff + summary) is computed server-side on save.

  const saveAll = useCallback((allSheets: any[]) => {
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
        // Backend computes the diff from DB; only the first sheet gets a changelog entry
        logChange: i === 0,
      });
    });
  }, [id]);

  const handleSaveNow = useCallback(() => {
    const all = latestSheetsRef.current || sheets;
    setSaveState('saving');
    saveAll(all);
    setIsDirty(false);
    // Work is now persisted server-side — drop the local draft and its restore banner.
    try { if (draftKey) localStorage.removeItem(draftKey); } catch { /* ignore */ }
    setDraftInfo(null);
    setSaveState('saved');
    setTimeout(() => setSaveState('idle'), 2000);
  }, [saveAll, sheets, draftKey]);

  const handleChange = useCallback((allSheets: any) => {
    if (!editor || !allSheets?.length) return;
    if (!acceptChangesRef.current) return;
    // First onChange after init timeout: treat as baseline (FortuneSheet may still be normalizing)
    if (!baselineTakenRef.current) {
      baselineTakenRef.current = true;
      latestSheetsRef.current = allSheets;
      lastSavedSheetsRef.current = JSON.parse(JSON.stringify(allSheets));
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

    // SSE is best-effort progress only. The POST below is the source of truth for completion,
    // so a dropped/buffered SSE connection can never leave the bar stuck.
    const evtSource = new EventSource(
      `/api/excel/${id}/import-progress?jobId=${jobId}&token=${encodeURIComponent(token || '')}`
    );
    evtSource.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        if (typeof data.progress === 'number') {
          setImportState((s) => (s.active ? { ...s, progress: Math.max(s.progress, data.progress) } : s));
        }
        if (data.done) evtSource.close();
      } catch { /* ignore parse errors */ }
    };
    evtSource.onerror = () => evtSource.close();

    const finishImport = async () => {
      const r = await api.get(`/spreadsheets/${id}`);
      const fresh = buildSheets(r.data);
      setSheets(fresh.length ? fresh : [{ name: 'Sheet1', index: 0, status: 1, celldata: [], config: {} }]);
      setWorkbookKey((k) => k + 1);
      setImportState({ active: false, progress: 0, error: null });
      setIsDirty(false);
    };

    try {
      const form = new FormData();
      form.append('file', file);
      // Resolves when the server has fully imported — this, not SSE, drives completion.
      await api.post(`/excel/${id}/import?jobId=${jobId}`, form);
      evtSource.close();
      await finishImport();
    } catch (err: any) {
      evtSource.close();
      setImportState({ active: false, progress: 0, error: err.response?.data?.error || 'Ошибка импорта' });
    }
  };

  // Stable hooks object — must NOT be recreated on re-renders, or FortuneSheet
  // re-registers hooks → infinite re-render loop. afterRenderCell fires for every
  // visible cell on each redraw; we record its exact CSS-pixel rect (relative to the
  // canvas top-left). FortuneSheet's canvas context is pre-scaled by devicePixelRatio,
  // so startX/startY/endX/endY are already CSS px — no ratio conversion.
  const workbookHooks = useMemo(() => ({
    // Broadcast our selection so other users see a live cursor. tabId is FortuneSheet's
    // (now stable) sheet id; r/c are the active cell of the selection range.
    afterSelectionChange: (tabId: string, sel: any) => {
      const r = sel?.row?.[0];
      const c = sel?.column?.[0];
      if (typeof r !== 'number' || typeof c !== 'number') return;
      socketRef.current?.emit('presence', { roomId: id, tabId, r, c });
    },
    afterRenderCell: (_cell: any, cellInfo: any, _ctx: any) => {
      const r = cellInfo.row as number;
      const c = cellInfo.column as number;
      pendingRectMapRef.current.set(`${r}_${c}`, {
        x: cellInfo.startX,
        y: cellInfo.startY,
        w: cellInfo.endX - cellInfo.startX,
        h: cellInfo.endY - cellInfo.startY,
      });
      // Coalesce a whole redraw pass (and continuous scrolling) into ONE update per animation
      // frame via rAF. A debounce would never commit during continuous scroll; an unthrottled
      // microtask fires many times per frame and floods React with renders (→ freeze/crash).
      if (renderBatchTimerRef.current === null) {
        renderBatchTimerRef.current = requestAnimationFrame(() => {
          cellRectMapRef.current = pendingRectMapRef.current;
          pendingRectMapRef.current = new Map();
          renderBatchTimerRef.current = null;
          // Only re-render when there are overlays to reposition (no cost on plain sheets)
          if (hasOverlaysRef.current) setMapVersion((v) => v + 1);
          // View changed → drop the stale cursor hover box and force re-detect on next move
          lastHoverKeyRef.current = null;
          setHoverOverlay((h) => (h ? null : h));
        }) as unknown as ReturnType<typeof setTimeout>;
      }
    },
  }), []); // eslint-disable-line react-hooks/exhaustive-deps

  // Memoize the heavy FortuneSheet element so per-frame overlay re-renders (mapVersion/hover)
  // don't re-render the whole grid. Only rebuilds when its real inputs change.
  const workbookEl = useMemo(() => (
    <Workbook
      key={workbookKey}
      ref={workbookRef}
      data={sheets}
      lang="ru"
      onChange={handleChange}
      showToolbar={editor}
      showFormulaBar
      allowEdit={editor}
      hooks={workbookHooks}
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
  ), [workbookKey, sheets, handleChange, editor, workbookHooks]);

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

      {/* Draft recovery banner — a leftover local draft means a previous session wasn't saved */}
      {editor && draftInfo && (
        <div className="bg-amber-50 border-b border-amber-200 px-4 py-2 flex items-center gap-2 text-sm text-amber-800 shrink-0 flex-wrap">
          <Info size={15} className="shrink-0" />
          <span>Найден несохранённый черновик от <strong>{formatTime(draftInfo.savedAt)}</strong>.</span>
          <button onClick={restoreDraft} className="ml-1 rounded-md border border-amber-300 bg-amber-100 px-2 py-0.5 text-xs font-medium hover:bg-amber-200 transition">Восстановить</button>
          <button onClick={discardDraft} className="rounded-md border border-amber-200 px-2 py-0.5 text-xs text-amber-600 hover:bg-amber-100 transition">Отклонить</button>
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
          onMouseMove={handleWorkbookMouseMove}
          onMouseDown={() => { setHoverOverlay(null); setNavHighlight(null); }}
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

          {/* Persistent overlays on every edited & currently-visible cell.
              Web-only decoration (absolutely-positioned divs) — never part of the sheet/export.
              Re-rendered whenever the cell map changes (mapVersion). */}
          {(() => {
            void mapVersion; // dependency: recompute positions after each redraw
            const origin = getCanvasOrigin();
            if (!origin) return null;
            const map = cellRectMapRef.current;
            const prefix = `${activeSheetIdx}_`;
            const seen = overlaySeenRef.current;
            const now = Date.now();
            const FADE_KEEP = 400; // ms to keep an off-screen overlay mounted so it can fade out

            // Refresh last-known positions for currently-visible edited cells
            for (const hKey in cellHighlights) {
              if (!hKey.startsWith(prefix)) continue;
              const rect = map.get(hKey.slice(prefix.length));
              if (rect) seen.set(hKey, { left: origin.left + rect.x, top: origin.top + rect.y, w: rect.w, h: rect.h, t: now });
            }

            const items: React.ReactNode[] = [];
            for (const [hKey, info] of seen) {
              const onScreen = map.has(hKey.slice(prefix.length)) && hKey.startsWith(prefix);
              if (!onScreen && now - info.t > FADE_KEEP) { seen.delete(hKey); continue; }
              const change = cellHighlights[hKey];
              if (!change) { seen.delete(hKey); continue; }
              const color = userColor(userColors, change.username);
              items.push(
                <div key={hKey} style={{
                  position: 'absolute', left: info.left, top: info.top, width: info.w, height: info.h,
                  background: color + '22', border: `1.5px solid ${color}`,
                  boxSizing: 'border-box', overflow: 'hidden',
                  pointerEvents: 'none', zIndex: 49,
                  opacity: onScreen ? 1 : 0,
                  transition: 'opacity 0.2s ease',
                }}>
                  {/* Date/time — inside the cell, top-right */}
                  <span style={{
                    position: 'absolute', top: 0, right: 0,
                    fontSize: 8, lineHeight: '10px', color: '#475569',
                    background: '#ffffffE6', borderLeft: '1px solid #e2e8f0',
                    borderBottom: '1px solid #e2e8f0', borderBottomLeftRadius: 3,
                    padding: '0 2px', whiteSpace: 'nowrap', maxWidth: '100%',
                    overflow: 'hidden', textOverflow: 'ellipsis',
                  }}>{formatTime(change.saved_at)}</span>
                  {/* Username — inside the cell, bottom-left */}
                  <span style={{
                    position: 'absolute', bottom: 0, left: 0,
                    fontSize: 9, lineHeight: '10px', color, padding: '0 2px 1px',
                    whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
                    maxWidth: '100%', fontWeight: 600,
                  }}>{change.username}</span>
                </div>
              );
            }
            return items;
          })()}

          {/* Navigation highlight — shown 2.5s after clicking a history entry */}
          {navHighlight && (
            <div
              style={{
                position: 'absolute',
                left:   navHighlight.left,
                top:    navHighlight.top,
                width:  navHighlight.width,
                height: navHighlight.height,
                background: navHighlight.color + '44',
                border: `2px solid ${navHighlight.color}`,
                pointerEvents: 'none',
                zIndex: 52,
                boxSizing: 'border-box',
                animation: 'twSheetPulse 0.6s ease-in-out 3',
              }}
            />
          )}

          {workbookEl}
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
                      const uColor = userColor(userColors, e.username);
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
