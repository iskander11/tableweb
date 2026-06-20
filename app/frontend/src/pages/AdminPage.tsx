import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { ArrowLeft, UserPlus, Trash2, ShieldCheck } from 'lucide-react';
import api from '../api/client';
import { useAuth } from '../store/auth';

interface User { id: string; username: string; email: string; role: string; is_active: boolean }

export default function AdminPage() {
  const navigate = useNavigate();
  const { user, logout } = useAuth();
  const qc = useQueryClient();

  const [form, setForm] = useState({ username: '', email: '', password: '', role: 'reader' });
  const [transferTo, setTransferTo] = useState('');

  const { data: users = [] } = useQuery<User[]>({
    queryKey: ['users'],
    queryFn: () => api.get('/auth/users').then((r) => r.data),
  });

  const createUser = useMutation({
    mutationFn: () => api.post('/auth/users', form),
    onSuccess: () => { qc.invalidateQueries({ queryKey: ['users'] }); setForm({ username: '', email: '', password: '', role: 'reader' }); },
  });

  const deleteUser = useMutation({
    mutationFn: (id: string) => api.delete(`/auth/users/${id}`),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['users'] }),
  });

  const changeRole = useMutation({
    mutationFn: ({ id, role }: { id: string; role: string }) => api.patch(`/auth/users/${id}/role`, { role }),
    onSuccess: () => qc.invalidateQueries({ queryKey: ['users'] }),
  });

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
            {(['username', 'email', 'password'] as const).map((field) => (
              <input
                key={field}
                type={field === 'password' ? 'password' : 'text'}
                placeholder={field === 'username' ? 'Логин' : field === 'email' ? 'Email' : 'Пароль'}
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
            onClick={() => createUser.mutate()}
            className="mt-4 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition text-sm"
          >
            Создать
          </button>
        </section>

        {/* User list */}
        <section className="bg-white rounded-xl border p-6">
          <h2 className="font-semibold text-gray-700 mb-4">Пользователи</h2>
          <div className="space-y-2">
            {users.filter((u) => u.is_active).map((u) => (
              <div key={u.id} className="flex items-center justify-between py-2 border-b last:border-0">
                <div>
                  <span className="font-medium text-gray-800">{u.username}</span>
                  <span className="text-xs text-gray-400 ml-2">{u.email}</span>
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
                      onClick={() => confirm(`Удалить ${u.username}?`) && deleteUser.mutate(u.id)}
                      className="p-1 text-red-400 hover:bg-red-50 rounded"
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
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
