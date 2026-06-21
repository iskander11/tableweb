// Allows the axios interceptor to trigger React Router navigation
// without a full page reload (which would reset the siteUnlocked state)
let _navigate: ((path: string) => void) | null = null;

export function registerNavigate(fn: (path: string) => void) {
  _navigate = fn;
}

export function navigateTo(path: string) {
  if (_navigate) {
    _navigate(path);
  } else {
    // Fallback only if React Router isn't mounted yet
    window.location.href = path;
  }
}
