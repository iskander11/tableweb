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

export default function SheetPage() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const { user, token, isEditor } = useAuth();
  const socketRef = useRef<Socket | null>(null);
  const [onlineUsers, setOnlineUsers] = useState<OnlineUser[]>([]);
  const [sheets, setSheets] = useState<any[]>([]);
  const workbookRef = useRef<any>(null);

  const { data: sheetMeta } = useQuery({
    queryKey: ['sheet', id],
    queryFn: () => api.get(`/spreadsheets/${id}`).then((r) => r.data),
  });

  useEffect(() => {
    if (!sheetMeta) return;

    // Initialize sheet data from DB
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
      // Apply remote changes to workbook
      workbookRef.current?.applyOp(changes);
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
    const form = new FormData();
    form.append('file', file);
    await api.post(`/excel/${id}/import`, form);
    window.location.reload();
  };

  if (!sheets.length) return <div className="flex items-center justify-center h-screen text-gray-400">Загрузка...</div>;

  return (
    <div className="flex flex-col h-screen bg-gray-50">
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
            <label className="flex items-center gap-1 cursor-pointer text-sm text-gray-600 hover:text-gray-800">
              <Upload size={15} /> Импорт
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport} />
            </label>
            <button onClick={handleExport} className="flex items-center gap-1 text-sm text-gray-600 hover:text-gray-800">
              <Download size={15} /> Экспорт
            </button>
          </>
        )}
      </header>

      <div className="flex-1 overflow-hidden">
        <Workbook
          ref={workbookRef}
          data={sheets}
          onChange={(data) => {
            if (isEditor()) {
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
  return Object.entries(cells).map(([key, val]) => {
    const [r, c] = key.split('_').map(Number);
    return { r, c, v: val };
  });
}

function unflattenCells(celldata: any[]) {
  const cells: Record<string, any> = {};
  celldata.forEach(({ r, c, v }) => { cells[`${r}_${c}`] = v; });
  return cells;
}
