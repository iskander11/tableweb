import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Plus, Pencil, Trash2, Lock, Archive } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';

interface Spreadsheet {
  id: string;
  name: string;
  creator_name: string;
  created_at: string;
  is_locked: boolean;
  backup_enabled: boolean;
}

export default function DashboardPage() {
  const { user, isAdmin, logout } = useAuth();
  const navigate = useNavigate();
  const qc = useQueryClient();
  const [newName, setNewName] = useState('');
  const [createError, setCreateError] = useState('');
  const [renameId, setRenameId] = useState<string | null>(null);
  const [renameName, setRenameName] = useState('');

  const { data: sheets = [] } = useQuery<Spreadsheet[]>({
    queryKey: ['sheets'],
    queryFn: () => api.get('/spreadsheets').then((r) => r.data),
  });

  const createMutation = useMutation({
    mutationFn: (name: string) => api.post('/spreadsheets', { name }),
    onSuccess: () => { qc.invalidateQueries({ queryKey: ['sheets'] }); setNewName(''); setCreateError(''); },
    onError: (e: any) => setCreateError(e.response?.data?.error || 'Не удалось создать таблицу'),
  });

  const deleteMutation = useMutation({
    mutationFn: (id: string) => api.delete(`/spreadsheets/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sheets'] }),
  });

  const renameMutation = useMutation({
    mutationFn: ({ id, name }: { id: string; name: string }) =>
      api.patch(`/spreadsheets/${id}/rename`, { name }),
    onSuccess: () => { qc.invalidateQueries({ queryKey: ['sheets'] }); setRenameId(null); },
    onError: (e: any) => alert(e.response?.data?.error || 'Не удалось переименовать'),
  });

  const toggleBackup = useMutation({
    mutationFn: (id: string) => api.patch(`/spreadsheets/${id}/backup-toggle`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sheets'] }),
  });

  const submitCreate = () => {
    const name = newName.trim();
    if (!name) { setCreateError('Введите название таблицы'); return; }
    createMutation.mutate(name);
  };

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b px-6 py-4 flex items-center justify-between">
        <h1 className="text-xl font-bold text-gray-800">TableWeb</h1>
        <div className="flex items-center gap-4">
          <span className="text-sm text-gray-600">{user?.username}</span>
          {isAdmin() && (
            <button onClick={() => navigate('/admin')} className="text-sm text-blue-600 hover:underline">
              Настройки
            </button>
          )}
          <button onClick={logout} className="text-sm text-red-500 hover:underline">Выйти</button>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8">
        <div className="flex items-center gap-3 mb-1">
          <input
            value={newName}
            onChange={(e) => { setNewName(e.target.value); if (createError) setCreateError(''); }}
            placeholder="Название новой таблицы..."
            className="border border-gray-300 rounded-lg px-3 py-2 flex-1 focus:outline-none focus:ring-2 focus:ring-blue-500"
            onKeyDown={(e) => e.key === 'Enter' && submitCreate()}
          />
          <button
            onClick={submitCreate}
            disabled={createMutation.isPending}
            className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 disabled:opacity-50 transition"
          >
            <Plus size={16} /> Создать
          </button>
        </div>
        {createError && <p className="text-red-500 text-sm mb-4">{createError}</p>}
        {!createError && <div className="mb-6" />}

        <div className="space-y-2">
          {sheets.map((sheet) => (
            <div
              key={sheet.id}
              className="bg-white rounded-xl border border-gray-200 px-5 py-4 flex items-center justify-between hover:border-blue-300 transition"
            >
              <div
                className="flex-1 cursor-pointer"
                onClick={() => !sheet.is_locked && navigate(`/sheet/${sheet.id}`)}
              >
                {renameId === sheet.id ? (
                  <input
                    value={renameName}
                    onChange={(e) => setRenameName(e.target.value)}
                    onBlur={() => renameMutation.mutate({ id: sheet.id, name: renameName })}
                    onKeyDown={(e) => e.key === 'Enter' && renameMutation.mutate({ id: sheet.id, name: renameName })}
                    className="border-b border-blue-500 outline-none text-gray-800 font-medium"
                    autoFocus
                    onClick={(e) => e.stopPropagation()}
                  />
                ) : (
                  <div className="flex items-center gap-2">
                    {sheet.is_locked && <Lock size={14} className="text-gray-400" />}
                    <span className="font-medium text-gray-800">{sheet.name}</span>
                    <span className="text-xs text-gray-400">· {sheet.creator_name}</span>
                  </div>
                )}
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => toggleBackup.mutate(sheet.id)}
                  className={`p-1.5 rounded transition ${sheet.backup_enabled ? 'text-green-600 bg-green-50 hover:bg-green-100' : 'text-gray-400 hover:bg-gray-100'}`}
                  title={sheet.backup_enabled ? 'Автобэкап включён (нажмите чтобы выключить)' : 'Включить автобэкап'}
                >
                  <Archive size={15} />
                </button>
                <button
                  onClick={() => { setRenameId(sheet.id); setRenameName(sheet.name); }}
                  className="p-1.5 rounded hover:bg-gray-100 text-gray-500"
                  title="Переименовать"
                >
                  <Pencil size={15} />
                </button>
                <button
                  onClick={() => confirm(`Удалить "${sheet.name}"?`) && deleteMutation.mutate(sheet.id)}
                  className="p-1.5 rounded hover:bg-red-50 text-red-400"
                  title="Удалить"
                >
                  <Trash2 size={15} />
                </button>
              </div>
            </div>
          ))}
          {sheets.length === 0 && (
            <p className="text-center text-gray-400 py-12">Нет таблиц. Создайте первую!</p>
          )}
        </div>
      </main>
    </div>
  );
}
