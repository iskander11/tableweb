import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { ArrowLeft, UserPlus, Trash2, ShieldCheck, Type, Upload, Archive, Download, RotateCcw } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';

interface User { id: string; username: string; email: string; role: string; is_active: boolean }
interface FontItem { id: string; displayName: string; familyName: string; format: string; url: string }
interface BackupItem { id: string; filename: string; created_at: string; size_bytes: number; username: string }

function formatBytes(n: number) {
  if (!n) return '0 Б';
  const u = ['Б', 'КБ', 'МБ', 'ГБ'];
  const i = Math.floor(Math.log(n) / Math.log(1024));
  return `${(n / Math.pow(1024, i)).toFixed(1)} ${u[i]}`;
}

export default function AdminPage() {
  const navigate = useNavigate();
  const { user, logout } = useAuth();
  const qc = useQueryClient();

  const [form, setForm] = useState({ username: '', password: '', role: 'reader' });
  const [transferTo, setTransferTo] = useState('');
  const [userError, setUserError] = useState('');
  const [fontFile, setFontFile] = useState<File | null>(null);
  const [fontName, setFontName] = useState('');
  const [fontError, setFontError] = useState('');

  const { data: users = [] } = useQuery<User[]>({
    queryKey: ['users'],
    queryFn: () => api.get('/auth/users').then((r) => r.data),
  });

  const createUser = useMutation({
    mutationFn: () => api.post('/auth/users', form),
    onSuccess: () => { qc.invalidateQueries({ queryKey: ['users'] }); setForm({ username: '', password: '', role: 'reader' }); setUserError(''); },
    onError: (e: any) => setUserError(e.response?.data?.error || 'Не удалось создать пользователя'),
  });

  const deleteUser = useMutation({
    mutationFn: (id: string) => api.delete(`/auth/users/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['users'] }),
  });

  const reactivateUser = useMutation({
    mutationFn: ({ id, role }: { id: string; role: string }) =>
      api.patch(`/auth/users/${id}/reactivate`, { role }),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['users'] }),
  });

  const changeRole = useMutation({
    mutationFn: ({ id, role }: { id: string; role: string }) => api.patch(`/auth/users/${id}/role`, { role }),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['users'] }),
  });

  const { data: fonts = [] } = useQuery<FontItem[]>({
    queryKey: ['fonts'],
    queryFn: () => api.get('/fonts').then((r) => r.data),
  });

  const uploadFont = useMutation({
    mutationFn: () => {
      const fd = new FormData();
      fd.append('file', fontFile as File);
      fd.append('displayName', fontName.trim());
      fd.append('familyName', fontName.trim());
      return api.post('/fonts', fd);
    },
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['fonts'] });
      setFontFile(null);
      setFontName('');
      setFontError('');
    },
    onError: (e: any) => setFontError(e.response?.data?.error || 'Ошибка загрузки шрифта'),
  });

  const deleteFont = useMutation({
    mutationFn: (id: string) => api.delete(`/fonts/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['fonts'] }),
  });

  const onFontFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    // Leave the name empty so the backend can auto-detect the real family name
    // from the file; the admin can still type a value to override.
    setFontFile(e.target.files?.[0] || null);
    setFontError('');
  };

  // --- Backups ---
  const { data: allSheets = [] } = useQuery<{ id: string }[]>({
    queryKey: ['sheets'],
    queryFn: () => api.get('/spreadsheets').then((r) => r.data),
  });
  const { data: backups = [] } = useQuery<BackupItem[]>({
    queryKey: ['backups'],
    queryFn: () => api.get('/backup').then((r) => r.data),
  });

  const createBackup = useMutation({
    mutationFn: () => api.post('/backup/all'),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['backups'] }),
    onError: (e: any) => alert(e.response?.data?.error || 'Не удалось создать бэкап'),
  });
  const restoreBackup = useMutation({
    mutationFn: (id: string) => api.post(`/backup/${id}/restore`),
    onSuccess: (r: any) => {
      qc.invalidateQueries({ queryKey: ['sheets'] });
      alert(`Восстановлено таблиц: ${r.data?.restored ?? 0}`);
    },
    onError: (e: any) => alert(e.response?.data?.error || 'Не удалось восстановить бэкап'),
  });
  const deleteBackup = useMutation({
    mutationFn: (id: string) => api.delete(`/backup/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['backups'] }),
  });

  const downloadBackup = async (b: BackupItem) => {
    const res = await api.get(`/backup/${b.id}/download`, { responseType: 'blob' });
    const url = URL.createObjectURL(res.data);
    const a = document.createElement('a');
    a.href = url;
    a.download = b.filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const transferAdmin = async () => {
    if (!transferTo || !confirm(`Передать права администратора пользователю ${transferTo}?`)) return;
    const target = users.find((u) => u.username === transferTo);
    if (!target) return alert('Пользователь не найден');
    await api.post('/auth/transfer-admin', { to_user_id: target.id });
    alert('Права переданы. Вы будете разлогинены.');
    logout();
    navigate('/login');
  };

  const otherUsers = users.filter((u) => u.id !== user?.id && u.is_active);

  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white border-b px-6 py-4 flex items-center gap-3">
        <button onClick={() => navigate('/')} className="p-1 rounded hover:bg-gray-100"><ArrowLeft size={18} /></button>
        <h1 className="font-bold text-gray-800 text-lg">Настройки администратора</h1>
      </header>

      <main className="max-w-3xl mx-auto px-6 py-8 space-y-8">
        {/* Create user */}
        <section className="bg-white rounded-xl border p-6">
          <h2 className="font-semibold text-gray-700 mb-4 flex items-center gap-2"><UserPlus size={16} /> Создать пользователя</h2>
          <div className="grid grid-cols-2 gap-3">
            {(['username', 'password'] as const).map((field) => (
              <input
                key={field}
                type={field === 'password' ? 'password' : 'text'}
                placeholder={field === 'username' ? 'Логин' : 'Пароль'}
                value={form[field]}
                onChange={(e) => setForm((f) => ({ ...f, [field]: e.target.value }))}
                className="border border-gray-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            ))}
            <select
              value={form.role}
              onChange={(e) => setForm((f) => ({ ...f, role: e.target.value }))}
              className="border border-gray-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="reader">Читатель</option>
              <option value="editor">Редактор</option>
            </select>
          </div>
          <button
            onClick={() => { setUserError(''); createUser.mutate(); }}
            disabled={createUser.isPending}
            className="mt-4 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 disabled:opacity-50 transition text-sm"
          >
            {createUser.isPending ? 'Создание…' : 'Создать'}
          </button>
          {userError && <p className="text-red-500 text-sm mt-2">{userError}</p>}
        </section>

        {/* User list */}
        <section className="bg-white rounded-xl border p-6">
          <h2 className="font-semibold text-gray-700 mb-4">Пользователи</h2>
          <div className="space-y-2">
            {users.filter((u) => u.is_active).map((u) => (
              <div key={u.id} className="flex items-center justify-between py-2 border-b last:border-0">
                <div>
                  <span className="font-medium text-gray-800">{u.username}</span>
                  <span className="text-xs text-gray-400 ml-2">{u.role}</span>
                </div>
                <div className="flex items-center gap-2">
                  <select
                    value={u.role}
                    onChange={(e) => changeRole.mutate({ id: u.id, role: e.target.value })}
                    disabled={u.id === user?.id}
                    className="text-sm border border-gray-200 rounded px-2 py-1"
                  >
                    <option value="reader">Читатель</option>
                    <option value="editor">Редактор</option>
                    <option value="admin">Администратор</option>
                  </select>
                  {u.id !== user?.id && (
                    <button
                      onClick={() => confirm(`Деактивировать ${u.username}?`) && deleteUser.mutate(u.id)}
                      className="p-1 text-red-400 hover:bg-red-50 rounded"
                      title="Деактивировать"
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>

          {/* Deactivated users */}
          {users.filter((u) => !u.is_active).length > 0 && (
            <div className="mt-4 pt-4 border-t">
              <p className="text-xs text-gray-400 font-medium mb-2 uppercase tracking-wide">Деактивированные</p>
              <div className="space-y-1">
                {users.filter((u) => !u.is_active).map((u) => (
                  <div key={u.id} className="flex items-center justify-between py-1.5">
                    <div>
                      <span className="text-gray-400 line-through text-sm">{u.username}</span>
                      <span className="text-xs text-gray-300 ml-2">{u.role}</span>
                    </div>
                    <button
                      onClick={() => reactivateUser.mutate({ id: u.id, role: u.role })}
                      className="text-xs text-blue-600 hover:bg-blue-50 border border-blue-200 rounded px-2 py-0.5"
                      title="Восстановить пользователя"
                    >
                      Восстановить
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>

        {/* Fonts */}
        <section className="bg-white rounded-xl border p-6">
          <h2 className="font-semibold text-gray-700 mb-1 flex items-center gap-2"><Type size={16} /> Шрифты</h2>
          <p className="text-sm text-gray-500 mb-4">
            Загрузите .ttf / .otf / .woff / .woff2 — <b>имя шрифта определится из файла автоматически</b>
            {' '}(то самое, что используется в Excel). При необходимости можно задать его вручную.
            Шрифт сразу применяется к импортированным таблицам и появляется в списке шрифтов на странице таблицы.
          </p>

          <div className="grid grid-cols-1 sm:grid-cols-[1fr_auto] gap-3 items-end">
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <label className="flex flex-col gap-1 text-xs text-gray-500">
                Файл шрифта
                <label className="flex items-center gap-2 cursor-pointer border border-gray-300 rounded-lg px-3 py-2 hover:bg-gray-50 text-sm text-gray-700 truncate">
                  <Upload size={14} className="shrink-0" />
                  <span className="truncate">{fontFile ? fontFile.name : 'Выбрать файл…'}</span>
                  <input
                    type="file"
                    accept=".ttf,.otf,.woff,.woff2"
                    className="hidden"
                    onChange={onFontFileChange}
                  />
                </label>
              </label>
              <label className="flex flex-col gap-1 text-xs text-gray-500">
                Имя шрифта (необязательно — определится из файла)
                <input
                  type="text"
                  placeholder="Авто (например, Times New Roman)"
                  value={fontName}
                  onChange={(e) => setFontName(e.target.value)}
                  className="border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </label>
            </div>
            <button
              onClick={() => uploadFont.mutate()}
              disabled={!fontFile || uploadFont.isPending}
              className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 disabled:opacity-50 transition text-sm h-[38px]"
            >
              {uploadFont.isPending ? 'Загрузка…' : 'Загрузить'}
            </button>
          </div>
          {fontError && <p className="text-red-500 text-xs mt-2">{fontError}</p>}

          <div className="mt-5 space-y-2">
            {fonts.length === 0 && <p className="text-sm text-gray-400">Пользовательских шрифтов пока нет.</p>}
            {fonts.map((f) => (
              <div key={f.id} className="flex items-center justify-between py-2 border-b last:border-0">
                <div className="min-w-0">
                  <span className="font-medium text-gray-800">{f.displayName}</span>
                  <span className="text-xs text-gray-400 ml-2 uppercase">{f.format}</span>
                  <div className="text-gray-600 truncate" style={{ fontFamily: `"${f.familyName}"` }}>
                    Пример текста · Sample 0123 · Привет
                  </div>
                </div>
                <button
                  onClick={() => confirm(`Удалить шрифт ${f.displayName}?`) && deleteFont.mutate(f.id)}
                  className="p-1 text-red-400 hover:bg-red-50 rounded shrink-0"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            ))}
          </div>
        </section>

        {/* Backups */}
        <section className="bg-white rounded-xl border p-6">
          <div className="flex items-center justify-between gap-3 flex-wrap mb-4">
            <h2 className="font-semibold text-gray-700 flex items-center gap-2 shrink-0"><Archive size={16} /> Резервные копии</h2>
            <button
              onClick={() => createBackup.mutate()}
              disabled={allSheets.length === 0 || createBackup.isPending}
              title={allSheets.length === 0 ? 'В системе нет таблиц' : 'Создать бэкап всех таблиц'}
              className="flex items-center gap-1.5 text-sm bg-green-600 text-white px-3 py-1.5 rounded-lg hover:bg-green-700 disabled:opacity-50 transition shrink-0"
            >
              <Archive size={14} /> {createBackup.isPending ? 'Создание…' : 'Создать бэкап'}
            </button>
          </div>

          {allSheets.length === 0 && (
            <p className="text-sm text-gray-400 mb-3">В системе нет таблиц — создание бэкапа недоступно.</p>
          )}

          <div className="space-y-2">
            {backups.length === 0 && <p className="text-sm text-gray-400">Бэкапов пока нет.</p>}
            {backups.map((b) => (
              <div key={b.id} className="flex items-center justify-between py-2 border-b last:border-0">
                <div className="min-w-0">
                  <div className="font-medium text-gray-800 truncate">{b.filename}</div>
                  <div className="text-xs text-gray-400">
                    {new Date(b.created_at).toLocaleString('ru-RU')} · {formatBytes(b.size_bytes)}
                    {b.username ? ` · ${b.username}` : ''}
                  </div>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => confirm('Восстановить таблицы из этого бэкапа в систему?') && restoreBackup.mutate(b.id)}
                    className="flex items-center gap-1 text-xs text-blue-600 hover:bg-blue-50 rounded px-2 py-1"
                    title="Вернуть таблицы в систему"
                  >
                    <RotateCcw size={14} /> Восстановить
                  </button>
                  <button onClick={() => downloadBackup(b)} className="p-1 text-gray-500 hover:bg-gray-100 rounded" title="Скачать">
                    <Download size={15} />
                  </button>
                  <button
                    onClick={() => confirm(`Удалить бэкап ${b.filename}?`) && deleteBackup.mutate(b.id)}
                    className="p-1 text-red-400 hover:bg-red-50 rounded"
                    title="Удалить"
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </section>

        {/* Transfer admin */}
        <section className="bg-white rounded-xl border p-6">
          <h2 className="font-semibold text-gray-700 mb-4 flex items-center gap-2"><ShieldCheck size={16} /> Передать права администратора</h2>
          <p className="text-sm text-gray-500 mb-3">При передаче прав вы станете читателем и будете разлогинены.</p>
          <div className="flex gap-3">
            <select
              value={transferTo}
              onChange={(e) => setTransferTo(e.target.value)}
              className="border border-gray-300 rounded-lg px-3 py-2 flex-1"
            >
              <option value="">Выберите пользователя</option>
              {otherUsers.map((u) => (
                <option key={u.id} value={u.username}>{u.username}</option>
              ))}
            </select>
            <button
              onClick={transferAdmin}
              disabled={!transferTo}
              className="bg-orange-500 text-white px-4 py-2 rounded-lg hover:bg-orange-600 disabled:opacity-50 transition text-sm"
            >
              Передать
            </button>
          </div>
        </section>
      </main>
    </div>
  );
}
