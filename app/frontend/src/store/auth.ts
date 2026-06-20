import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import api from '../api/client';

interface User {
  id: string;
  username: string;
  role: 'admin' | 'editor' | 'reader';
}

interface AuthStore {
  user: User | null;
  token: string | null;
  login: (username: string, password: string) => Promise<void>;
  logout: () => void;
  isAdmin: () => boolean;
  isEditor: () => boolean;
}

export const useAuth = create<AuthStore>()(
  persist(
    (set, get) => ({
      user: null,
      token: null,
      login: async (username, password) => {
        const { data } = await api.post('/auth/login', { username, password });
        localStorage.setItem('token', data.token);
        set({ user: data.user, token: data.token });
      },
      logout: () => {
        localStorage.removeItem('token');
        set({ user: null, token: null });
      },
      isAdmin: () => get().user?.role === 'admin',
      isEditor: () => ['admin', 'editor'].includes(get().user?.role ?? ''),
    }),
    { name: 'auth-store', partialize: (s) => ({ user: s.user, token: s.token }) }
  )
);
