import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Plus, Pencil, Trash2, Lock, Archive, ChevronLeft, ChevronRight, KeyRound, X } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';

function ChangePasswordModal({ onClose }: { onClose: () => void }) {
  const [form, setForm] = useState({ current: '', next: '', confirm: '' });
  const [error, setError] = useState('');
  const [success, setSuccess] = useState(false);

  const mutation = useMutation({
    mutationFn: () => api.patch('/auth/password', { currentPassword: form.current, newPassword: form.next }),
    onSuccess: () => { setSuccess(true); setError(''); },
    onError: (e: any) => setError(e.response?.data?.error || 'Не удалось сменить пароль'),
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    if (form.next !== form.confirm) { setError('Новые пароли не совпадают'); return; }
    if (form.next.length < 4) { setError('Пароль должен быть не менее 4 символов'); return; }
    mutation.mutate();
  };

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center px-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm p-6">
        <div className="flex items-center justify-between mb-5">
          <h2 className="font-semibold text-gray-800 flex items-center gap-2">
            <KeyRound size={16} /> Смена пароля
          </h2>
          <button onClick={onClose} className="p-1 rounded hover:bg-gray-100 text-gray-400">
            <X size={18} />
          </button>
        </div>

        {success ? (
          <div className="text-center py-4">
            <p className="text-green-600 font-medium mb-4">Пароль успешно изменён</p>
            <button onClick={onClose} className="bg-blue-600 text-white px-5 py-2 rounded-lg text-sm hover:bg-blue-700 transition">
              Закрыть
            </button>
          </div>
        ) : (
          <form onSubmit={handleSubmit} className="space-y-3">
            <input
              type="password"
              placeholder="Текущий пароль"
              value={form.current}
              onChange={(e) => setForm((f) => ({ ...f, current: e.target.value }))}
              className="w-full border border-gray-300 rounded-lg px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500 text-base"
              required
            />
            <input
              type="password"
              placeholder="Новый пароль"
              value={form.next}
              onChange={(e) => setForm((f) => ({ ...f, next: e.target.value }))}
              className="w-full border border-gray-300 rounded-lg px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500 text-base"
              required
            />
            <input
              type="password"
              placeholder="Повторите новый пароль"
              value={form.confirm}
              onChange={(e) => setForm((f) => ({ ...f, confirm: e.target.value }))}
              className="w-full border border-gray-300 rounded-lg px-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500 text-base"
              required
            />
            {error && <p className="text-red-500 text-sm">{error}</p>}
            <button
              type="submit"
              disabled={mutation.isPending}
              className="w-full bg-blue-600 text-white rounded-lg py-2.5 font-medium hover:bg-blue-700 disabled:opacity-50 transition"
            >
              {mutation.isPending ? 'Сохранение…' : 'Сменить пароль'}
            </button>
          </form>
        )}
      </div>
    </div>
  );
}

interface Spreadsheet {
  id: string;
  name: string;
  creator_name: string;
  created_at: string;
  is_locked: boolean;
  backup_enabled: boolean;
}

const PAGE_SIZE = 15;

export default function DashboardPage() {
  const { user, isAdmin, logout } = useAuth();
  const navigate = useNavigate();
  const qc = useQueryClient();
  const [newName, setNewName] = useState('');
  const [createError, setCreateError] = useState('');
  const [renameId, setRenameId] = useState<string | null>(null);
  const [renameName, setRenameName] = useState('');
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState('');
  const [showChangePassword, setShowChangePassword] = useState(false);

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

  const filtered = sheets.filter((s) =>
    s.name.toLowerCase().includes(search.toLowerCase())
  );
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  const safePage = Math.min(page, totalPages);
  const pageItems = filtered.slice((safePage - 1) * PAGE_SIZE, safePage * PAGE_SIZE);

  const handleSearch = (v: string) => { setSearch(v); setPage(1); };

  return (
    <div className="min-h-screen bg-gray-50">
      {showChangePassword && <ChangePasswordModal onClose={() => setShowChangePassword(false)} />}

      <header className="bg-white border-b px-4 sm:px-6 py-4 flex items-center justify-between">
        <h1 className="text-xl font-bold text-gray-800">TableWeb</h1>
        <div className="flex items-center gap-3 sm:gap-4">
          <button
            onClick={() => setShowChangePassword(true)}
            className="flex items-center gap-1.5 text-sm text-gray-600 hover:text-gray-800 hidden sm:flex"
            title="Сменить пароль"
          >
            <span className="hidden md:inline">{user?.username}</span>
            <KeyRound size={14} />
          </button>
          <button
            onClick={() => setShowChangePassword(true)}
            className="p-1.5 rounded hover:bg-gray-100 text-gray-500 sm:hidden"
            title="Сменить пароль"
          >
            <KeyRound size={16} />
          </button>
          {isAdmin() && (
            <button onClick={() => navigate('/admin')} className="text-sm text-blue-600 hover:underline">
              Настройки
            </button>
          )}
          <button onClick={logout} className="text-sm text-red-500 hover:underline">Выйти</button>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 sm:px-6 py-8">
        {/* Create row */}
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
            className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 disabled:opacity-50 transition shrink-0"
          >
            <Plus size={16} /> <span className="hidden sm:inline">Создать</span>
          </button>
        </div>
        {createError && <p className="text-red-500 text-sm mb-4">{createError}</p>}
        {!createError && <div className="mb-4" />}

        {/* Search — shown only if more than PAGE_SIZE tables */}
        {sheets.length > PAGE_SIZE && (
          <input
            value={search}
            onChange={(e) => handleSearch(e.target.value)}
            placeholder="Поиск по названию..."
            className="w-full border border-gray-200 rounded-lg px-3 py-2 mb-4 focus:outline-none focus:ring-2 focus:ring-blue-400 text-sm"
          />
        )}

        <div className="space-y-2">
          {pageItems.map((sheet) => (
            <div
              key={sheet.id}
              className="bg-white rounded-xl border border-gray-200 px-4 sm:px-5 py-4 flex items-center justify-between hover:border-blue-300 transition"
            >
              <div
                className="flex-1 min-w-0 cursor-pointer"
                onClick={() => !sheet.is_locked && navigate(`/sheet/${sheet.id}`)}
              >
                {renameId === sheet.id ? (
                  <input
                    value={renameName}
                    onChange={(e) => setRenameName(e.target.value)}
                    onBlur={() => renameMutation.mutate({ id: sheet.id, name: renameName })}
                    onKeyDown={(e) => e.key === 'Enter' && renameMutation.mutate({ id: sheet.id, name: renameName })}
                    className="border-b border-blue-500 outline-none text-gray-800 font-medium w-full"
                    autoFocus
                    onClick={(e) => e.stopPropagation()}
                  />
                ) : (
                  <div className="flex items-center gap-2 min-w-0">
                    {sheet.is_locked && <Lock size={14} className="text-gray-400 shrink-0" />}
                    <span className="font-medium text-gray-800 truncate">{sheet.name}</span>
                    <span className="text-xs text-gray-400 shrink-0 hidden sm:inline">· {sheet.creator_name}</span>
                  </div>
                )}
              </div>
              <div className="flex items-center gap-1.5 shrink-0 ml-2">
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

          {filtered.length === 0 && search && (
            <p className="text-center text-gray-400 py-8">Ничего не найдено по запросу «{search}»</p>
          )}
          {sheets.length === 0 && (
            <p className="text-center text-gray-400 py-12">Нет таблиц. Создайте первую!</p>
          )}
        </div>

        {/* Pagination */}
        {totalPages > 1 && (
          <div className="flex items-center justify-center gap-2 mt-6">
            <button
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              disabled={safePage === 1}
              className="p-2 rounded-lg border border-gray-200 hover:bg-gray-50 disabled:opacity-40 transition"
            >
              <ChevronLeft size={16} />
            </button>
            {Array.from({ length: totalPages }, (_, i) => i + 1)
              .filter((p) => p === 1 || p === totalPages || Math.abs(p - safePage) <= 2)
              .reduce<(number | '...')[]>((acc, p, i, arr) => {
                if (i > 0 && typeof arr[i - 1] === 'number' && (p as number) - (arr[i - 1] as number) > 1) acc.push('...');
                acc.push(p);
                return acc;
              }, [])
              .map((p, i) =>
                p === '...' ? (
                  <span key={`e${i}`} className="px-2 text-gray-400 text-sm">…</span>
                ) : (
                  <button
                    key={p}
                    onClick={() => setPage(p as number)}
                    className={`w-9 h-9 rounded-lg text-sm font-medium transition ${
                      safePage === p
                        ? 'bg-blue-600 text-white'
                        : 'border border-gray-200 hover:bg-gray-50 text-gray-700'
                    }`}
                  >
                    {p}
                  </button>
                )
              )}
            <button
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              disabled={safePage === totalPages}
              className="p-2 rounded-lg border border-gray-200 hover:bg-gray-50 disabled:opacity-40 transition"
            >
              <ChevronRight size={16} />
            </button>
            <span className="text-xs text-gray-400 ml-1">
              {filtered.length} таблиц
            </span>
          </div>
        )}
      </main>
    </div>
  );
}
