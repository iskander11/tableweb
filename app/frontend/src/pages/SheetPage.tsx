import { useEffect, useRef, useState, useCallback } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import { io, Socket } from 'socket.io-client';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import { ArrowLeft, Download, Upload, Users, Save } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';
import { useFonts } from '../fonts/useFonts';

interface OnlineUser { id: string; username: string }

interface ImportState {
  active: boolean;
  progress: number;
  error: string | null;
}

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
  const [saveState, setSaveState] = useState<'idle' | 'saved'>('idle');
  const workbookRef = useRef<any>(null);
  const latestSheetsRef = useRef<any[] | null>(null);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const { version: fontsVersion } = useFonts();

  useEffect(() => {
    if (fontsVersion > 0) setWorkbookKey((k) => k + 1);
  }, [fontsVersion]);

  const { data: sheetMeta, isError: sheetError } = useQuery({
    queryKey: ['sheet', id],
    queryFn: () => api.get(`/spreadsheets/${id}`).then((r) => r.data),
    retry: false,
  });

  useEffect(() => {
    if (!sheetMeta) return;
    const built = buildSheets(sheetMeta);
    setSheets(built.length ? built : [{ name: 'Sheet1', index: 0, status: 1, celldata: [], config: {} }]);
  }, [sheetMeta]);

  useEffect(() => {
    if (!token) return;
    const socket = io('/', { auth: { token } });
    socketRef.current = socket;
    socket.emit('join-sheet', id);
    socket.on('room-users', setOnlineUsers);
    socket.on('cell-change', ({ changes }) => {
      workbookRef.current?.applyOp?.(changes);
    });
    return () => { socket.disconnect(); };
  }, [id, token]);

  const saveSheet = useCallback((data: any, index: number) => {
    if (!data) return;
    socketRef.current?.emit('save-sheet', {
      sheetId: id,
      sheetIndex: index,
      data: {
        name: data.name,
        cells: cellsFromSheet(data),
        columnWidths: data.config?.columnlen || {},
        rowHeights: data.config?.rowlen || {},
        merges: data.config?.merge || {},
        borderInfo: data.config?.borderInfo || [],
        filterSelect: data.filter_select || null,
        filterCriteria: data.filter || null,
        frozen: data.frozen || null,
      },
    });
  }, [id]);

  const saveAll = useCallback((allSheets: any[]) => {
    (allSheets || []).forEach((s, i) => saveSheet(s, s.index ?? i));
  }, [saveSheet]);

  const handleSaveNow = useCallback(() => {
    const all = latestSheetsRef.current || sheets;
    saveAll(all);
    setSaveState('saved');
    setTimeout(() => setSaveState('idle'), 1800);
  }, [saveAll, sheets]);

  const handleChange = useCallback((allSheets: any) => {
    if (!isEditor() || !allSheets?.length) return;
    latestSheetsRef.current = allSheets;
    socketRef.current?.emit('cell-change', { sheetId: id, changes: allSheets });
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => saveAll(allSheets), 2000);
  }, [id, isEditor, saveAll]);

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

      <header className="bg-white border-b px-4 py-2 flex items-center gap-3 shrink-0">
        <button onClick={() => navigate('/')} className="p-1 rounded hover:bg-gray-100 shrink-0">
          <ArrowLeft size={18} />
        </button>
        <h2 className="font-semibold text-gray-800 flex-1 min-w-0 truncate">{sheetMeta?.name}</h2>

        <div className="flex items-center gap-2 shrink-0">
          {onlineUsers.length > 0 && (
            <div className="flex items-center gap-1 text-xs text-gray-500 max-w-[180px] overflow-hidden">
              <Users size={13} className="shrink-0" />
              {onlineUsers.map((u) => (
                <span key={u.id} className="bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded truncate">{u.username}</span>
              ))}
            </div>
          )}

          {isEditor() && (
            <>
              <button onClick={handleSaveNow} title="Сохранить все изменения"
                className={`flex items-center gap-1.5 text-sm rounded-lg px-3 py-1.5 border transition ${
                  saveState === 'saved'
                    ? 'border-green-200 bg-green-50 text-green-600'
                    : 'border-gray-200 text-gray-600 hover:text-gray-800 hover:bg-gray-50'
                }`}>
                <Save size={14} /> {saveState === 'saved' ? 'Сохранено' : 'Сохранить'}
              </button>
              <label title="Импорт .xlsx" className="flex items-center gap-1.5 cursor-pointer text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-3 py-1.5 hover:bg-gray-50 transition">
                <Upload size={14} /> Импорт
                <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport} />
              </label>
              <button onClick={handleExport} title="Экспорт .xlsx"
                className="flex items-center gap-1.5 text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-3 py-1.5 hover:bg-gray-50 transition">
                <Download size={14} /> Экспорт
              </button>
            </>
          )}
        </div>
      </header>

      <div className="flex-1 overflow-hidden">
        <Workbook
          key={workbookKey}
          ref={workbookRef}
          data={sheets}
          lang="ru"
          onChange={handleChange}
          showToolbar={isEditor()}
          showFormulaBar
          allowEdit={isEditor()}
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
    </div>
  );
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
