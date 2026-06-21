import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { useEffect, useState } from 'react';
import { useAuth } from './store/auth';
import LoginPage from './pages/LoginPage';
import DashboardPage from './pages/DashboardPage';
import SheetPage from './pages/SheetPage';
import AdminPage from './pages/AdminPage';
import SiteGatePage from './pages/SiteGatePage';
import api from './api/client';

const qc = new QueryClient();

function ProtectedRoute({ children }: { children: React.ReactNode }) {
  const { user } = useAuth();
  return user ? <>{children}</> : <Navigate to="/login" replace />;
}

function AdminRoute({ children }: { children: React.ReactNode }) {
  const { user } = useAuth();
  if (!user) return <Navigate to="/login" replace />;
  if (user.role !== 'admin') return <Navigate to="/" replace />;
  return <>{children}</>;
}

function AppContent() {
  // null = checking, true = unlocked, false = need site password
  const [siteUnlocked, setSiteUnlocked] = useState<boolean | null>(null);

  useEffect(() => {
    // Try a lightweight request to check if site cookie is valid
    api.get('/auth/me').then(() => {
      setSiteUnlocked(true);
    }).catch((err) => {
      if (err.response?.data?.requireSiteAuth) {
        setSiteUnlocked(false);
      } else {
        // 401 from JWT (not logged in) — site is unlocked, just need user login
        setSiteUnlocked(true);
      }
    });
  }, []);

  if (siteUnlocked === null) {
    return (
      <div className="min-h-screen bg-gray-900 flex items-center justify-center">
        <div className="w-8 h-8 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (siteUnlocked === false) {
    return <SiteGatePage onUnlocked={() => setSiteUnlocked(true)} />;
  }

  return (
    <BrowserRouter>
      <Routes>
        <Route path="/login" element={<LoginPage />} />
        <Route path="/" element={<ProtectedRoute><DashboardPage /></ProtectedRoute>} />
        <Route path="/sheet/:id" element={<ProtectedRoute><SheetPage /></ProtectedRoute>} />
        <Route path="/admin" element={<AdminRoute><AdminPage /></AdminRoute>} />
      </Routes>
    </BrowserRouter>
  );
}

export default function App() {
  return (
    <QueryClientProvider client={qc}>
      <AppContent />
    </QueryClientProvider>
  );
}
