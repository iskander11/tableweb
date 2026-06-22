import { Router } from 'express';
import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import { query } from '../db/index.js';
import { authenticate, requireAdmin } from '../middleware/auth.js';

const PRESET_COLORS = [
  '#EF4444','#F97316','#EAB308','#22C55E','#14B8A6',
  '#3B82F6','#8B5CF6','#EC4899','#06B6D4','#84CC16',
];
function randomColor() { return PRESET_COLORS[Math.floor(Math.random() * PRESET_COLORS.length)]; }

const router = Router();

// Login
router.post('/login', async (req, res) => {
  const { username, password } = req.body;
  try {
    const { rows } = await query(
      'SELECT * FROM users WHERE username = $1 AND is_active = TRUE',
      [username]
    );
    const user = rows[0];
    if (!user || !(await bcrypt.compare(password, user.password_hash))) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }
    const token = jwt.sign(
      { id: user.id, username: user.username, role: user.role },
      process.env.JWT_SECRET,
      { expiresIn: '7d' }
    );
    res.json({ token, user: { id: user.id, username: user.username, role: user.role } });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create user (admin only)
router.post('/users', authenticate, requireAdmin, async (req, res) => {
  const { username, password, role } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: 'Логин и пароль обязательны' });
  }
  if (username.length < 2 || username.length > 50) {
    return res.status(400).json({ error: 'Логин должен быть от 2 до 50 символов' });
  }
  if (password.length < 4) {
    return res.status(400).json({ error: 'Пароль должен быть не менее 4 символов' });
  }
  try {
    const trimmed = username.trim();
    const existing = await query(
      'SELECT id, is_active FROM users WHERE LOWER(username) = LOWER($1)',
      [trimmed]
    );
    if (existing.rows.length) {
      const found = existing.rows[0];
      if (found.is_active) {
        return res.status(409).json({ error: `Пользователь «${trimmed}» уже существует` });
      }
      // Reactivate previously deleted user
      const hash = await bcrypt.hash(password, 12);
      const { rows } = await query(
        `UPDATE users SET is_active = TRUE, password_hash = $1, role = $2, created_by = $3
         WHERE id = $4 RETURNING id, username, role`,
        [hash, role || 'reader', req.user.id, found.id]
      );
      return res.status(201).json(rows[0]);
    }
    const hash = await bcrypt.hash(password, 12);
    const { rows } = await query(
      'INSERT INTO users (username, email, password_hash, role, created_by, color) VALUES ($1,$2,$3,$4,$5,$6) RETURNING id, username, role, color',
      [trimmed, `${trimmed}@tableweb.local`, hash, role || 'reader', req.user.id, randomColor()]
    );
    res.status(201).json(rows[0]);
  } catch (err) {
    if (err.code === '23505') {
      return res.status(409).json({ error: `Пользователь «${username.trim()}» уже существует` });
    }
    res.status(500).json({ error: 'Внутренняя ошибка сервера' });
  }
});

// Get all users (admin only)
router.get('/users', authenticate, requireAdmin, async (req, res) => {
  const { rows } = await query(
    'SELECT id, username, email, role, created_at, is_active, color FROM users ORDER BY created_at'
  );
  res.json(rows);
});

// Get username→color map for ALL users (incl. deactivated — historical edits still
// need their author's color). Accessible to all authenticated users.
router.get('/user-colors', authenticate, async (req, res) => {
  const { rows } = await query(
    'SELECT username, color FROM users'
  );
  res.json(rows);
});

// Change own color
router.patch('/me/color', authenticate, async (req, res) => {
  const { color } = req.body;
  if (!color || !/^#[0-9A-Fa-f]{6}$/.test(color)) {
    return res.status(400).json({ error: 'Неверный формат цвета' });
  }
  await query('UPDATE users SET color = $1 WHERE id = $2', [color, req.user.id]);
  res.json({ success: true, color });
});

// Update user role (admin only)
router.patch('/users/:id/role', authenticate, requireAdmin, async (req, res) => {
  const { role } = req.body;
  await query('UPDATE users SET role = $1 WHERE id = $2', [role, req.params.id]);
  res.json({ success: true });
});

// Deactivate user (admin only)
router.delete('/users/:id', authenticate, requireAdmin, async (req, res) => {
  await query('UPDATE users SET is_active = FALSE WHERE id = $1', [req.params.id]);
  res.json({ success: true });
});

// Reactivate user (admin only)
router.patch('/users/:id/reactivate', authenticate, requireAdmin, async (req, res) => {
  const { role } = req.body;
  await query(
    'UPDATE users SET is_active = TRUE, role = COALESCE($1, role) WHERE id = $2',
    [role || null, req.params.id]
  );
  res.json({ success: true });
});

// Permanently delete user (admin only) — only allowed if user is inactive
router.delete('/users/:id/permanent', authenticate, requireAdmin, async (req, res) => {
  try {
    const { rows } = await query('SELECT is_active FROM users WHERE id = $1', [req.params.id]);
    if (!rows.length) return res.status(404).json({ error: 'Пользователь не найден' });
    if (rows[0].is_active) return res.status(400).json({ error: 'Нельзя удалить активного пользователя' });
    await query('DELETE FROM users WHERE id = $1', [req.params.id]);
    res.json({ success: true });
  } catch (err) {
    if (err.code === '23503') {
      return res.status(409).json({ error: 'Нельзя удалить: у пользователя есть связанные данные (таблицы, история)' });
    }
    res.status(500).json({ error: err.message });
  }
});

// Transfer admin role
router.post('/transfer-admin', authenticate, requireAdmin, async (req, res) => {
  const { to_user_id } = req.body;
  try {
    await query('UPDATE users SET role = $1 WHERE id = $2', ['reader', req.user.id]);
    await query('UPDATE users SET role = $1 WHERE id = $2', ['admin', to_user_id]);
    await query(
      'INSERT INTO admin_transfers (from_user_id, to_user_id) VALUES ($1, $2)',
      [req.user.id, to_user_id]
    );
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Change own password
router.patch('/password', authenticate, async (req, res) => {
  const { currentPassword, newPassword } = req.body;
  if (!currentPassword || !newPassword) {
    return res.status(400).json({ error: 'Укажите текущий и новый пароль' });
  }
  if (newPassword.length < 4) {
    return res.status(400).json({ error: 'Новый пароль должен быть не менее 4 символов' });
  }
  try {
    const { rows } = await query('SELECT password_hash FROM users WHERE id = $1', [req.user.id]);
    if (!rows.length) return res.status(404).json({ error: 'Пользователь не найден' });
    const valid = await bcrypt.compare(currentPassword, rows[0].password_hash);
    if (!valid) return res.status(401).json({ error: 'Неверный текущий пароль' });
    const hash = await bcrypt.hash(newPassword, 12);
    await query('UPDATE users SET password_hash = $1 WHERE id = $2', [hash, req.user.id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get current user
router.get('/me', authenticate, async (req, res) => {
  const { rows } = await query(
    'SELECT id, username, email, role FROM users WHERE id = $1',
    [req.user.id]
  );
  res.json(rows[0]);
});

export default router;
