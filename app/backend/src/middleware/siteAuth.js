import crypto from 'crypto';

const COOKIE_NAME = 'site_token';
const COOKIE_MAX_AGE = 7 * 24 * 60 * 60 * 1000; // 7 days

function makeToken(password) {
  return crypto.createHmac('sha256', process.env.JWT_SECRET || 'fallback')
    .update(password)
    .digest('hex');
}

// POST /api/site-auth — validate site password, set cookie
export function siteAuthLogin(req, res) {
  const { password } = req.body;
  const sitePass = process.env.SITE_PASSWORD;

  if (!sitePass) {
    // No site password configured — site is open
    return res.json({ ok: true });
  }

  if (!password || password !== sitePass) {
    return res.status(401).json({ error: 'Неверный пароль доступа' });
  }

  const token = makeToken(password);
  res.cookie(COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: 'lax',
    maxAge: COOKIE_MAX_AGE,
    path: '/',
  });
  res.json({ ok: true });
}

// Middleware — check cookie on every /api/* request (except /api/site-auth itself)
export function requireSiteAuth(req, res, next) {
  const sitePass = process.env.SITE_PASSWORD;

  // Site password not configured → skip
  if (!sitePass) return next();

  // Allow the site-auth login endpoint through
  if (req.path === '/api/site-auth') return next();

  const token = req.cookies?.[COOKIE_NAME];
  const expected = makeToken(sitePass);

  if (!token || token !== expected) {
    return res.status(401).json({ requireSiteAuth: true });
  }

  next();
}
