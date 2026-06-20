import { Router } from 'express';
import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import { query } from '../db/index.js';
import { authenticate, requireAdmin } from '../middleware/auth.js';

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
  try {
    const hash = await bcrypt.hash(password, 12);
    const { rows } = await query(
      'INSERT INTO users (username, email, password_hash, role, created_by) VALUES ($1,$2,$3,$4,$5) RETURNING id, username, role',
      [username, `${username}@tableweb.local`, hash, role || 'reader', req.user.id]
    );
    res.status(201).json(rows[0]);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// Get all users (admin only)
router.get('/users', authenticate, requireAdmin, async (req, res) => {
  const { rows } = await query(
    'SELECT id, username, email, role, created_at, is_active FROM users ORDER BY created_at'
  );
  res.json(rows);
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

// Get current user
router.get('/me', authenticate, async (req, res) => {
  const { rows } = await query(
    'SELECT id, username, email, role FROM users WHERE id = $1',
    [req.user.id]
  );
  res.json(rows[0]);
});

export default router;
