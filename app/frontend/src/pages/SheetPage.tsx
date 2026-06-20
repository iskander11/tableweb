import { useEffect, useRef, useState, useCallback } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { useQuery } from '@tanstack/react-query';
import { io, Socket } from 'socket.io-client';
import { Workbook } from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import { ArrowLeft, Download, Upload, Users } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';

interface OnlineUser { id: string; username: string }

interface ImportState {
  active: boolean;
  progress: number;
  error: string | null;
}

export default function SheetPage() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const { user, token, isEditor } = useAuth();
  const socketRef = useRef<Socket | null>(null);
  const [onlineUsers, setOnlineUsers] = useState<OnlineUser[]>([]);
  const [sheets, setSheets] = useState<any[]>([]);
  const [importState, setImportState] = useState<ImportState>({ active: false, progress: 0, error: null });
  const workbookRef = useRef<any>(null);

  const { data: sheetMeta } = useQuery({
    queryKey: ['sheet', id],
    queryFn: () => api.get(`/spreadsheets/${id}`).then((r) => r.data),
  });

  useEffect(() => {
    if (!sheetMeta) return;

    const initialSheets = (sheetMeta.sheets || []).map((s: any, i: number) => ({
      name: s.data?.name || `Sheet${i + 1}`,
      index: i,
      status: 1,
      celldata: flattenCells(s.data?.cells || {}),
      config: {
        columnlen: s.data?.columnWidths || {},
        rowlen: s.data?.rowHeights || {},
        merge: s.data?.merges || {},
      },
    }));
    setSheets(initialSheets.length ? initialSheets : [{ name: 'Sheet1', index: 0, status: 1, celldata: [] }]);
  }, [sheetMeta]);

  useEffect(() => {
    const socket = io('/', { auth: { token } });
    socketRef.current = socket;

    socket.emit('join-sheet', id);
    socket.on('room-users', setOnlineUsers);
    socket.on('cell-change', ({ changes }) => {
      workbookRef.current?.applyOp?.(changes);
    });

    return () => { socket.disconnect(); };
  }, [id, token]);

  const handleCellChange = useCallback((op: any) => {
    socketRef.current?.emit('cell-change', { sheetId: id, changes: op });
  }, [id]);

  const handleSave = useCallback((data: any) => {
    socketRef.current?.emit('save-sheet', {
      sheetId: id,
      sheetIndex: data.index || 0,
      data: {
        name: data.name,
        cells: unflattenCells(data.celldata || []),
        columnWidths: data.config?.columnlen || {},
        rowHeights: data.config?.rowlen || {},
        merges: data.config?.merge || {},
      },
    });
  }, [id]);

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
    setImportState({ active: true, progress: 0, error: null });

    // Subscribe to progress via SSE
    const evtSource = new EventSource(`/api/excel/${id}/import-progress?jobId=${jobId}`, {});
    evtSource.onmessage = (event) => {
      const data = JSON.parse(event.data);
      setImportState((s) => ({ ...s, progress: data.progress }));
      if (data.done) {
        evtSource.close();
        if (data.error) {
          setImportState({ active: false, progress: 0, error: data.error });
        } else {
          // Reload page to get fresh data from server
          setTimeout(() => window.location.reload(), 300);
        }
      }
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

  if (!sheets.length) {
    return (
      <div className="flex items-center justify-center h-screen text-gray-400">
        Загрузка таблицы...
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-gray-50">
      {/* Import progress overlay */}
      {importState.active && (
        <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center">
          <div className="bg-white rounded-2xl shadow-xl p-8 w-80">
            <h3 className="font-semibold text-gray-800 mb-4 text-center">Импорт файла</h3>
            <div className="w-full bg-gray-100 rounded-full h-3 mb-3">
              <div
                className="bg-blue-500 h-3 rounded-full transition-all duration-300"
                style={{ width: `${importState.progress}%` }}
              />
            </div>
            <p className="text-center text-sm text-gray-500">
              {importState.progress < 100 ? `${importState.progress}% — обработка...` : 'Сохранение...'}
            </p>
          </div>
        </div>
      )}

      {/* Error toast */}
      {importState.error && (
        <div className="fixed top-4 right-4 z-50 bg-red-50 border border-red-200 rounded-xl px-4 py-3 shadow">
          <p className="text-red-600 text-sm">{importState.error}</p>
          <button
            onClick={() => setImportState((s) => ({ ...s, error: null }))}
            className="text-xs text-red-400 underline mt-1"
          >
            Закрыть
          </button>
        </div>
      )}

      <header className="bg-white border-b px-4 py-2 flex items-center gap-3">
        <button onClick={() => navigate('/')} className="p-1 rounded hover:bg-gray-100">
          <ArrowLeft size={18} />
        </button>
        <h2 className="font-semibold text-gray-800 flex-1">{sheetMeta?.name}</h2>

        <div className="flex items-center gap-1 text-xs text-gray-500">
          <Users size={13} />
          {onlineUsers.map((u) => (
            <span key={u.id} className="bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded">{u.username}</span>
          ))}
        </div>

        {isEditor() && (
          <>
            <label className="flex items-center gap-1 cursor-pointer text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-3 py-1.5 hover:bg-gray-50">
              <Upload size={14} /> Импорт .xlsx
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport} />
            </label>
            <button
              onClick={handleExport}
              className="flex items-center gap-1 text-sm text-gray-600 hover:text-gray-800 border border-gray-200 rounded-lg px-3 py-1.5 hover:bg-gray-50"
            >
              <Download size={14} /> Экспорт .xlsx
            </button>
          </>
        )}
      </header>

      <div className="flex-1 overflow-hidden">
        <Workbook
          ref={workbookRef}
          data={sheets}
          onChange={(data: any) => {
            if (isEditor() && data?.length) {
              handleCellChange(data);
              handleSave(data[0]);
            }
          }}
          showToolbar={isEditor()}
          showFormulaBar={true}
          allowEdit={isEditor()}
        />
      </div>
    </div>
  );
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
    if (v !== undefined) cells[`${r}_${c}`] = v;
  });
  return cells;
}
