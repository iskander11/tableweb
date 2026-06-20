import { Router } from 'express';
import archiver from 'archiver';
import { exportExcel } from '../services/excel.js';
import { query } from '../db/index.js';
import { authenticate, requireAdmin } from '../middleware/auth.js';
import { createWriteStream, mkdirSync } from 'fs';
import { join } from 'path';

const router = Router();
const BACKUP_DIR = process.env.BACKUP_DIR || './backups';
mkdirSync(BACKUP_DIR, { recursive: true });

// Manual backup of all spreadsheets
router.post('/all', authenticate, requireAdmin, async (req, res) => {
  try {
    const { rows: sheets } = await query('SELECT * FROM spreadsheets');
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
        'INSERT INTO backups (filename, created_by, size_bytes) VALUES ($1, $2, $3)',
        [filename, req.user.id, size]
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

export default router;
