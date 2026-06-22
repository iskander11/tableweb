import { Router } from 'express';
import multer from 'multer';
import { importExcel, exportExcel } from '../services/excel.js';
import pool, { query } from '../db/index.js';
import { authenticate } from '../middleware/auth.js';

const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// Import progress via SSE — client subscribes before uploading
// GET /api/excel/:id/import-progress?jobId=xxx
const activeJobs = new Map(); // jobId -> { progress, done, error, clients[] }

router.get('/:id/import-progress', authenticate, (req, res) => {
  const { jobId } = req.query;
  if (!jobId) return res.status(400).json({ error: 'jobId required' });

  res.set({
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
    // Tell nginx not to buffer this response, so progress events reach the client
    // immediately and the bar animates smoothly instead of jumping 2%→100% at the end.
    'X-Accel-Buffering': 'no',
  });
  res.flushHeaders();

  const sendEvent = (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);

  if (!activeJobs.has(jobId)) {
    activeJobs.set(jobId, { progress: 0, done: false, error: null, clients: [] });
  }
  const job = activeJobs.get(jobId);
  job.clients.push(sendEvent);

  // Send current progress immediately
  sendEvent({ progress: job.progress, done: job.done, error: job.error });

  req.on('close', () => {
    const j = activeJobs.get(jobId);
    if (j) j.clients = j.clients.filter((c) => c !== sendEvent);
  });
});

// POST /api/excel/:id/import?jobId=xxx
router.post('/:id/import', authenticate, upload.single('file'), async (req, res) => {
  const { jobId } = req.query;

  const updateProgress = (pct) => {
    if (!jobId) return;
    const job = activeJobs.get(jobId);
    if (!job) return;
    job.progress = pct;
    job.clients.forEach((send) => send({ progress: pct, done: false }));
  };

  try {
    if (jobId && !activeJobs.has(jobId)) {
      activeJobs.set(jobId, { progress: 0, done: false, error: null, clients: [] });
    }

    updateProgress(5);
    const sheets = await importExcel(req.file.buffer, updateProgress);
    updateProgress(90);

    // Atomic swap: DELETE + INSERTs in one transaction so a mid-import failure
    // can never leave the table partially wiped — either the new data fully lands
    // or the old data stays untouched (ROLLBACK).
    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      await client.query('DELETE FROM spreadsheet_data WHERE spreadsheet_id = $1', [req.params.id]);
      for (let i = 0; i < sheets.length; i++) {
        await client.query(
          'INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data) VALUES ($1, $2, $3)',
          [req.params.id, i, JSON.stringify(sheets[i])]
        );
      }
      await client.query('UPDATE spreadsheets SET updated_at = NOW() WHERE id = $1', [req.params.id]);
      await client.query('COMMIT');
    } catch (txErr) {
      await client.query('ROLLBACK').catch(() => {});
      throw txErr;
    } finally {
      client.release();
    }

    updateProgress(100);
    if (jobId) {
      const job = activeJobs.get(jobId);
      if (job) {
        job.done = true;
        job.clients.forEach((send) => send({ progress: 100, done: true }));
        // cleanup after 30s
        setTimeout(() => activeJobs.delete(jobId), 30000);
      }
    }

    res.json({ success: true, sheets: sheets.length });
  } catch (err) {
    if (jobId) {
      const job = activeJobs.get(jobId);
      if (job) {
        job.error = err.message;
        job.clients.forEach((send) => send({ progress: 0, done: true, error: err.message }));
      }
    }
    res.status(500).json({ error: err.message });
  }
});

// Export to Excel
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
