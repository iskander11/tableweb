import { Router } from 'express';
import { query } from '../db/index.js';
import { authenticate, requireAdmin } from '../middleware/auth.js';

const router = Router();

// Get spreadsheets accessible to the user
router.get('/', authenticate, async (req, res) => {
  try {
    let rows;
    if (req.user.role === 'admin') {
      ({ rows } = await query(
        `SELECT s.*, u.username as creator_name FROM spreadsheets s
         JOIN users u ON s.created_by = u.id ORDER BY s.created_at DESC`
      ));
    } else {
      ({ rows } = await query(
        `SELECT s.*, u.username as creator_name, sp.role as my_role FROM spreadsheets s
         JOIN users u ON s.created_by = u.id
         JOIN spreadsheet_permissions sp ON sp.spreadsheet_id = s.id AND sp.user_id = $1
         ORDER BY s.created_at DESC`,
        [req.user.id]
      ));
    }
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Create spreadsheet
router.post('/', authenticate, async (req, res) => {
  const name = (req.body.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Введите название таблицы' });
  try {
    const dup = await query('SELECT 1 FROM spreadsheets WHERE LOWER(name) = LOWER($1)', [name]);
    if (dup.rows.length) {
      return res.status(409).json({ error: 'Таблица с таким названием уже существует' });
    }
    const { rows } = await query(
      'INSERT INTO spreadsheets (name, created_by) VALUES ($1, $2) RETURNING *',
      [name, req.user.id]
    );
    const sheet = rows[0];
    // Initialize empty data
    await query(
      'INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data) VALUES ($1, 0, $2)',
      [sheet.id, JSON.stringify({ cells: {}, columnWidths: {}, rowHeights: {} })]
    );
    res.status(201).json(sheet);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get single spreadsheet with data
router.get('/:id', authenticate, async (req, res) => {
  try {
    const { rows: [sheet] } = await query(
      'SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]
    );
    if (!sheet) return res.status(404).json({ error: 'Not found' });

    const canAccess = req.user.role === 'admin' ||
      sheet.created_by === req.user.id;

    if (!canAccess) {
      const { rows } = await query(
        'SELECT role FROM spreadsheet_permissions WHERE spreadsheet_id = $1 AND user_id = $2',
        [req.params.id, req.user.id]
      );
      if (!rows.length) return res.status(403).json({ error: 'No access' });
    }

    const { rows: dataRows } = await query(
      'SELECT * FROM spreadsheet_data WHERE spreadsheet_id = $1 ORDER BY sheet_index',
      [req.params.id]
    );

    const { rows: permissions } = await query(
      `SELECT sp.*, u.username FROM spreadsheet_permissions sp
       JOIN users u ON sp.user_id = u.id WHERE sp.spreadsheet_id = $1`,
      [req.params.id]
    );

    res.json({ ...sheet, sheets: dataRows, permissions });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Rename spreadsheet
router.patch('/:id/rename', authenticate, async (req, res) => {
  const name = (req.body.name || '').trim();
  if (!name) return res.status(400).json({ error: 'Введите название таблицы' });
  const { rows: [sheet] } = await query('SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]);
  if (!sheet) return res.status(404).json({ error: 'Not found' });
  if (sheet.created_by !== req.user.id && req.user.role !== 'admin')
    return res.status(403).json({ error: 'Forbidden' });

  const dup = await query(
    'SELECT 1 FROM spreadsheets WHERE LOWER(name) = LOWER($1) AND id <> $2',
    [name, req.params.id]
  );
  if (dup.rows.length) {
    return res.status(409).json({ error: 'Таблица с таким названием уже существует' });
  }

  await query('UPDATE spreadsheets SET name = $1, updated_at = NOW() WHERE id = $2', [name, req.params.id]);
  res.json({ success: true });
});

// Delete spreadsheet
router.delete('/:id', authenticate, async (req, res) => {
  const { rows: [sheet] } = await query('SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]);
  if (!sheet) return res.status(404).json({ error: 'Not found' });
  if (sheet.created_by !== req.user.id && req.user.role !== 'admin')
    return res.status(403).json({ error: 'Forbidden' });

  await query('DELETE FROM spreadsheets WHERE id = $1', [req.params.id]);
  res.json({ success: true });
});

// Set permissions on spreadsheet
router.post('/:id/permissions', authenticate, async (req, res) => {
  const { user_id, role } = req.body;
  const { rows: [sheet] } = await query('SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]);
  if (sheet.created_by !== req.user.id && req.user.role !== 'admin')
    return res.status(403).json({ error: 'Forbidden' });

  await query(
    `INSERT INTO spreadsheet_permissions (spreadsheet_id, user_id, role)
     VALUES ($1, $2, $3)
     ON CONFLICT (spreadsheet_id, user_id) DO UPDATE SET role = $3`,
    [req.params.id, user_id, role]
  );
  res.json({ success: true });
});

// Remove permission
router.delete('/:id/permissions/:userId', authenticate, async (req, res) => {
  const { rows: [sheet] } = await query('SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]);
  if (sheet.created_by !== req.user.id && req.user.role !== 'admin')
    return res.status(403).json({ error: 'Forbidden' });

  await query(
    'DELETE FROM spreadsheet_permissions WHERE spreadsheet_id = $1 AND user_id = $2',
    [req.params.id, req.params.userId]
  );
  res.json({ success: true });
});

// Toggle lock
router.patch('/:id/lock', authenticate, requireAdmin, async (req, res) => {
  await query(
    'UPDATE spreadsheets SET is_locked = NOT is_locked WHERE id = $1',
    [req.params.id]
  );
  res.json({ success: true });
});

// Get change log for a spreadsheet
router.get('/:id/changelog', authenticate, async (req, res) => {
  try {
    const { rows } = await query(
      `SELECT username, sheet_index, summary, saved_at
       FROM change_log WHERE spreadsheet_id = $1
       ORDER BY saved_at DESC LIMIT 100`,
      [req.params.id]
    );
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Toggle backup
router.patch('/:id/backup-toggle', authenticate, async (req, res) => {
  const { rows: [sheet] } = await query('SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]);
  if (sheet.created_by !== req.user.id && req.user.role !== 'admin')
    return res.status(403).json({ error: 'Forbidden' });

  await query(
    'UPDATE spreadsheets SET backup_enabled = NOT backup_enabled WHERE id = $1',
    [req.params.id]
  );
  res.json({ success: true });
});

export default router;
