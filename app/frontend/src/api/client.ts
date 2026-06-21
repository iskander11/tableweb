import axios from 'axios';
import { navigateTo } from './navigationService';

const api = axios.create({
  baseURL: '/api',
  withCredentials: true,  // always send cookies (needed for site-token cookie)
});

api.interceptors.request.use((config) => {
  const token = localStorage.getItem('token');
  if (token) config.headers.Authorization = `Bearer ${token}`;
  return config;
});

api.interceptors.response.use(
  (r) => r,
  (err) => {
    if (err.response?.status === 401) {
      // Site password required — do NOT redirect, let App.tsx show the gate
      if (err.response?.data?.requireSiteAuth) {
        return Promise.reject(err);
      }
      // Expired/invalid JWT — clear token and navigate to login WITHOUT page reload
      if (localStorage.getItem('token')) {
        localStorage.removeItem('token');
        navigateTo('/login');
      }
    }
    return Promise.reject(err);
  }
);

export default api;
