import axios from 'axios';

const api = axios.create({ baseURL: '/api' });

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) config.headers.Authorization = `Bearer ${token}`;
  return config;
});

api.interceptors.response.use(
  (r) => r,
  (err) => {
    if (err.response?.status === 401) {
      // Site password required — let App.tsx handle it, do not redirect
      if (err.response?.data?.requireSiteAuth) {
        return Promise.reject(err);
      }
      // Expired/invalid JWT — only redirect if we actually had a token
      if (localStorage.getItem('token')) {
        localStorage.removeItem('token');
        window.location.href = '/login';
      }
    }
    return Promise.reject(err);
  }
);

export default api;
