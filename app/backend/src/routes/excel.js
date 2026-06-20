import { Router } from 'express';
import multer from 'multer';
import { importExcel, exportExcel } from '../services/excel.js';
import { query } from '../db/index.js';
import { authenticate } from '../middleware/auth.js';

const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// Import Excel into existing spreadsheet
router.post('/:id/import', authenticate, upload.single('file'), async (req, res) => {
  try {
    const sheets = await importExcel(req.file.buffer);

    // Clear existing sheets data and replace with imported
    await query('DELETE FROM spreadsheet_data WHERE spreadsheet_id = $1', [req.params.id]);

    for (let i = 0; i < sheets.length; i++) {
      await query(
        'INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data) VALUES ($1, $2, $3)',
        [req.params.id, i, JSON.stringify(sheets[i])]
      );
    }
    await query('UPDATE spreadsheets SET updated_at = NOW() WHERE id = $1', [req.params.id]);

    res.json({ success: true, sheets: sheets.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Export spreadsheet to Excel
router.get('/:id/export', authenticate, async (req, res) => {
  try {
    const { rows: dataRows } = await query(
      'SELECT data FROM spreadsheet_data WHERE spreadsheet_id = $1 ORDER BY sheet_index',
      [req.params.id]
    );
    const { rows: [sheet] } = await query('SELECT name FROM spreadsheets WHERE id = $1', [req.params.id]);

    const sheetsData = dataRows.map((r) => r.data);
    const buffer = await exportExcel(sheetsData);

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${encodeURIComponent(sheet.name)}.xlsx"`,
    });
    res.send(Buffer.from(buffer));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default router;
