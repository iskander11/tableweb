import { BrowserRouter, Routes, Route, Navigate, useNavigate } from 'react-router-dom';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { useEffect, useState } from 'react';
import { useAuth } from './store/auth';
import { registerNavigate } from './api/navigationService';
import LoginPage from './pages/LoginPage';
import DashboardPage from './pages/DashboardPage';
import SheetPage from './pages/SheetPage';
import AdminPage from './pages/AdminPage';
import SiteGatePage from './pages/SiteGatePage';
import api from './api/client';

const qc = new QueryClient();

// Registers React Router's navigate with the axios interceptor
// so 401 JWT errors redirect without a page reload
function NavigationRegistrar() {
  const navigate = useNavigate();
  useEffect(() => {
    registerNavigate(navigate);
  }, [navigate]);
  return null;
}

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

const SESSION_KEY = 'site_unlocked';

function AppContent() {
  const [siteUnlocked, setSiteUnlocked] = useState<boolean | null>(() => {
    // Check sessionStorage first to avoid flicker on React re-renders
    return sessionStorage.getItem(SESSION_KEY) === '1' ? true : null;
  });

  useEffect(() => {
    // Already know we're unlocked (from sessionStorage) — verify in background
    if (siteUnlocked === true) {
      api.get('/site-ping').catch((err) => {
        if (err.response?.data?.requireSiteAuth) {
          sessionStorage.removeItem(SESSION_KEY);
          setSiteUnlocked(false);
        }
      });
      return;
    }

    // Unknown state — check with server
    api.get('/site-ping').then(() => {
      sessionStorage.setItem(SESSION_KEY, '1');
      setSiteUnlocked(true);
    }).catch((err) => {
      if (err.response?.data?.requireSiteAuth) {
        setSiteUnlocked(false);
      } else {
        // Network error or other — assume unlocked (site ping endpoint may not exist)
        setSiteUnlocked(true);
      }
    });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const handleUnlocked = () => {
    sessionStorage.setItem(SESSION_KEY, '1');
    setSiteUnlocked(true);
  };

  if (siteUnlocked === null) {
    return (
      <div className="min-h-screen bg-gray-900 flex items-center justify-center">
        <div className="w-8 h-8 border-2 border-blue-500 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (siteUnlocked === false) {
    return <SiteGatePage onUnlocked={handleUnlocked} />;
  }

  return (
    <BrowserRouter>
      <NavigationRegistrar />
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
