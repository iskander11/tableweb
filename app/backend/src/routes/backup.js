import { Router } from 'express';
import archiver from 'archiver';
import AdmZip from 'adm-zip';
import { exportExcel, importExcel } from '../services/excel.js';
import { query } from '../db/index.js';
import { authenticate, requireAdmin } from '../middleware/auth.js';
import { canManageSheet } from '../services/permissions.js';
import { createWriteStream, mkdirSync, existsSync, unlinkSync } from 'fs';
import { join } from 'path';

const router = Router();
const BACKUP_DIR = process.env.BACKUP_DIR || './backups';
mkdirSync(BACKUP_DIR, { recursive: true });

// Manual backup of all spreadsheets
router.post('/all', authenticate, requireAdmin, async (req, res) => {
  try {
    const { rows: sheets } = await query('SELECT * FROM spreadsheets');
    if (sheets.length === 0) {
      return res.status(400).json({ error: 'В системе нет таблиц — бэкап недоступен' });
    }
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `backup-all-${timestamp}.zip`;
    const filepath = join(BACKUP_DIR, filename);

    const output = createWriteStream(filepath);
    const archive = archiver('zip');
    archive.pipe(output);

    for (const sheet of sheets) {
      const { rows: dataRows } = await query(
        'SELECT data FROM spreadsheet_data WHERE spreadsheet_id = $1 ORDER BY sheet_index',
        [sheet.id]
      );
      const sheetsData = dataRows.map((r) => r.data);
      const buffer = await exportExcel(sheetsData);
      archive.append(Buffer.from(buffer), { name: `${sheet.name}.xlsx` });
    }

    await archive.finalize();

    output.on('close', async () => {
      const size = archive.pointer();
      await query(
        'INSERT INTO backups (filename, created_by, size_bytes, backup_type) VALUES ($1, $2, $3, $4)',
        [filename, req.user.id, size, 'manual']
      );
      res.json({ success: true, filename, size });
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Daily backup of a single spreadsheet (admins on any table, editors on their own)
router.post('/sheet/:id', authenticate, async (req, res) => {
  try {
    const { rows: [sheet] } = await query('SELECT * FROM spreadsheets WHERE id = $1', [req.params.id]);
    if (!sheet) return res.status(404).json({ error: 'Таблица не найдена' });
    if (!(await canManageSheet(req.user, req.params.id))) {
      return res.status(403).json({ error: 'Нет прав на резервную копию этой таблицы' });
    }

    const { rows: dataRows } = await query(
      'SELECT data FROM spreadsheet_data WHERE spreadsheet_id = $1 ORDER BY sheet_index',
      [req.params.id]
    );
    const buffer = await exportExcel(dataRows.map((r) => r.data));

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const safeName = sheet.name.replace(/[/\\?%*:|"<>]/g, '_');
    const filename = `daily-${safeName}-${timestamp}.zip`;
    const filepath = join(BACKUP_DIR, filename);

    const output = createWriteStream(filepath);
    const archive = archiver('zip');
    archive.pipe(output);
    archive.append(Buffer.from(buffer), { name: `${sheet.name}.xlsx` });
    await archive.finalize();

    output.on('close', async () => {
      const size = archive.pointer();
      await query(
        'INSERT INTO backups (spreadsheet_id, filename, created_by, size_bytes, backup_type, sheet_name) VALUES ($1,$2,$3,$4,$5,$6)',
        [sheet.id, filename, req.user.id, size, 'daily', sheet.name]
      );
      res.json({ success: true, filename, size });
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// List backups
router.get('/', authenticate, requireAdmin, async (req, res) => {
  const { rows } = await query(
    `SELECT b.*, u.username FROM backups b
     LEFT JOIN users u ON b.created_by = u.id ORDER BY b.created_at DESC`
  );
  res.json(rows);
});

// Download a backup zip
router.get('/:id/download', authenticate, requireAdmin, async (req, res) => {
  const { rows: [bk] } = await query('SELECT * FROM backups WHERE id = $1', [req.params.id]);
  if (!bk) return res.status(404).json({ error: 'Бэкап не найден' });
  const filepath = join(BACKUP_DIR, bk.filename);
  if (!existsSync(filepath)) return res.status(404).json({ error: 'Файл бэкапа отсутствует на диске' });
  res.download(filepath, bk.filename);
});

// Delete a backup (file + record)
router.delete('/:id', authenticate, requireAdmin, async (req, res) => {
  const { rows: [bk] } = await query('SELECT * FROM backups WHERE id = $1', [req.params.id]);
  if (bk) {
    const filepath = join(BACKUP_DIR, bk.filename);
    if (existsSync(filepath)) { try { unlinkSync(filepath); } catch { /* ignore */ } }
  }
  await query('DELETE FROM backups WHERE id = $1', [req.params.id]);
  res.json({ success: true });
});

// Restore a backup into the system: each .xlsx becomes a spreadsheet
router.post('/:id/restore', authenticate, requireAdmin, async (req, res) => {
  try {
    const { rows: [bk] } = await query('SELECT * FROM backups WHERE id = $1', [req.params.id]);
    if (!bk) return res.status(404).json({ error: 'Бэкап не найден' });
    const filepath = join(BACKUP_DIR, bk.filename);
    if (!existsSync(filepath)) return res.status(404).json({ error: 'Файл бэкапа отсутствует на диске' });

    const zip = new AdmZip(filepath);
    const entries = zip.getEntries().filter((e) => !e.isDirectory && /\.xlsx$/i.test(e.entryName));
    if (entries.length === 0) return res.status(400).json({ error: 'В бэкапе нет таблиц' });

    let restored = 0;
    for (const entry of entries) {
      const baseName = entry.entryName.replace(/\.xlsx$/i, '').split('/').pop();

      // Avoid duplicate names: append a suffix if needed
      let name = baseName;
      let n = 1;
      // eslint-disable-next-line no-await-in-loop
      while ((await query('SELECT 1 FROM spreadsheets WHERE LOWER(name) = LOWER($1)', [name])).rows.length) {
        name = `${baseName} (восстановлено${n > 1 ? ' ' + n : ''})`;
        n += 1;
      }

      // eslint-disable-next-line no-await-in-loop
      const sheets = await importExcel(entry.getData());
      // eslint-disable-next-line no-await-in-loop
      const { rows: [created] } = await query(
        'INSERT INTO spreadsheets (name, created_by) VALUES ($1, $2) RETURNING id',
        [name, req.user.id]
      );
      for (let i = 0; i < sheets.length; i += 1) {
        // eslint-disable-next-line no-await-in-loop
        await query(
          'INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data) VALUES ($1, $2, $3)',
          [created.id, i, JSON.stringify(sheets[i])]
        );
      }
      restored += 1;
    }

    res.json({ success: true, restored });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default router;
