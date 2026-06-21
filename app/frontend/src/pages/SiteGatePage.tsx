import { useState } from 'react';
import { Lock } from 'lucide-react';
import api from '../api/client';

interface Props {
  onUnlocked: () => void;
}

export default function SiteGatePage({ onUnlocked }: Props) {
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      await api.post('/site-auth', { password });
      onUnlocked();
    } catch {
      setError('Неверный пароль доступа');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-900 flex items-center justify-center px-4">
      <div className="bg-gray-800 rounded-2xl shadow-2xl p-8 w-full max-w-sm border border-gray-700">
        <div className="flex justify-center mb-5">
          <div className="w-14 h-14 rounded-full bg-blue-600/20 border border-blue-500/30 flex items-center justify-center">
            <Lock size={24} className="text-blue-400" />
          </div>
        </div>
        <h1 className="text-xl font-bold text-white text-center mb-1">TableWeb</h1>
        <p className="text-gray-400 text-sm text-center mb-6">Введите пароль доступа к системе</p>
        <form onSubmit={handleSubmit} className="space-y-4">
          <input
            type="password"
            value={password}
            onChange={(e) => { setPassword(e.target.value); setError(''); }}
            placeholder="Пароль доступа"
            autoFocus
            className="w-full bg-gray-700 border border-gray-600 text-white placeholder-gray-500 rounded-lg px-4 py-3 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-base"
            required
          />
          {error && (
            <p className="text-red-400 text-sm text-center">{error}</p>
          )}
          <button
            type="submit"
            disabled={loading || !password}
            className="w-full bg-blue-600 text-white rounded-lg py-3 font-medium hover:bg-blue-700 disabled:opacity-50 transition text-base"
          >
            {loading ? 'Проверка...' : 'Войти'}
          </button>
        </form>
      </div>
    </div>
  );
}
