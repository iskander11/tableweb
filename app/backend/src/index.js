import express from 'express';
import { createServer } from 'http';
import { Server } from 'socket.io';
import cors from 'cors';
import compression from 'compression';
import cookieParser from 'cookie-parser';
import dotenv from 'dotenv';
import cron from 'node-cron';
import jwt from 'jsonwebtoken';

import { siteAuthLogin, requireSiteAuth } from './middleware/siteAuth.js';
import authRoutes from './routes/auth.js';
import spreadsheetRoutes from './routes/spreadsheets.js';
import excelRoutes from './routes/excel.js';
import backupRoutes from './routes/backup.js';
import fontRoutes, { ensureFontsTable, FONTS_DIR } from './routes/fonts.js';
import { query } from './db/index.js';
import { exportExcel } from './services/excel.js';
import { createWriteStream, mkdirSync } from 'fs';
import { join } from 'path';
import archiver from 'archiver';

dotenv.config();

// Max change_log entries retained per spreadsheet (older ones are pruned on each save).
const CHANGELOG_LIMIT = Number(process.env.CHANGELOG_LIMIT) || 200;

const app = express();
const httpServer = createServer(app);

const io = new Server(httpServer, {
  cors: { origin: process.env.FRONTEND_URL || '*', credentials: true },
});

app.use(cors({ origin: process.env.FRONTEND_URL || '*', credentials: true }));
app.use(compression());
app.use(cookieParser());
app.use(express.json({ limit: '50mb' }));

// Site-level password gate (before all other routes)
app.post('/api/site-auth', siteAuthLogin);
app.use(requireSiteAuth);

// Lightweight ping — returns 200 if site cookie is valid (no JWT needed)
app.get('/api/site-ping', (_req, res) => res.json({ ok: true }));

app.use('/api/auth', authRoutes);
app.use('/api/spreadsheets', spreadsheetRoutes);
app.use('/api/excel', excelRoutes);
app.use('/api/backup', backupRoutes);
// Serve font files without auth so the browser font loader can fetch them.
// Registered before the /api/fonts router so the static path takes priority.
app.use('/api/fonts/files', express.static(FONTS_DIR, { maxAge: '7d', immutable: true }));
app.use('/api/fonts', fontRoutes);

ensureFontsTable().catch((err) => console.error('fonts table init failed:', err));

// Track online users per spreadsheet room
const roomUsers = {};

io.use((socket, next) => {
  const token = socket.handshake.auth.token;
  try {
    socket.user = jwt.verify(token, process.env.JWT_SECRET);
    next();
  } catch {
    next(new Error('Unauthorized'));
  }
});

io.on('connection', (socket) => {
  socket.on('join-sheet', (sheetId) => {
    socket.join(sheetId);
    if (!roomUsers[sheetId]) roomUsers[sheetId] = {};
    roomUsers[sheetId][socket.id] = { id: socket.user.id, username: socket.user.username };
    io.to(sheetId).emit('room-users', Object.values(roomUsers[sheetId]));
  });

  socket.on('cell-change', ({ sheetId, changes }) => {
    socket.to(sheetId).emit('cell-change', { userId: socket.user.id, changes });
  });

  // Broadcast color change to all connected clients
  socket.on('update-color', ({ color }) => {
    io.emit('user-color-changed', { username: socket.user.username, color });
  });

  socket.on('save-sheet', async ({ sheetId, sheetIndex, data, logChange }) => {
    try {
      // Compute diff server-side before overwriting (backend always has the ground truth)
      let changedCells = null;
      let summary = null;
      if (logChange) {
        try {
          const prev = await query(
            'SELECT data FROM spreadsheet_data WHERE spreadsheet_id = $1 AND sheet_index = $2',
            [sheetId, sheetIndex]
          );
          if (prev.rows.length > 0) {
            const diff = computeSheetDiff(prev.rows[0].data, data);
            changedCells = diff.cells.length > 0 ? diff.cells : null;
            summary = diff.summary;
          }
        } catch (_) { /* diff failure must not block the save */ }
      }

      await query(
        `INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data)
         VALUES ($1, $2, $3)
         ON CONFLICT (spreadsheet_id, sheet_index) DO UPDATE SET data = $3, updated_at = NOW()`,
        [sheetId, sheetIndex, JSON.stringify(data)]
      );
      await query('UPDATE spreadsheets SET updated_at = NOW() WHERE id = $1', [sheetId]);

      if (logChange) {
        await query(
          `INSERT INTO change_log (spreadsheet_id, sheet_index, user_id, username, summary, changed_cells)
           VALUES ($1, $2, $3, $4, $5, $6)`,
          [sheetId, sheetIndex, socket.user.id, socket.user.username, summary || null,
           changedCells ? JSON.stringify(changedCells) : null]
        );
        // Retention: keep only the newest CHANGELOG_LIMIT entries per spreadsheet so the
        // history table can't grow unbounded and slow down queries over time.
        await query(
          `DELETE FROM change_log
           WHERE spreadsheet_id = $1
             AND id NOT IN (
               SELECT id FROM change_log WHERE spreadsheet_id = $1
               ORDER BY saved_at DESC LIMIT $2
             )`,
          [sheetId, CHANGELOG_LIMIT]
        ).catch(() => { /* retention is best-effort; never block a save */ });
        const entry = {
          username: socket.user.username,
          sheet_index: sheetIndex,
          summary: summary || null,
          changed_cells: changedCells || null,
          saved_at: new Date().toISOString(),
        };
        io.to(sheetId).emit('changelog-update', entry);
      }
    } catch (err) {
      socket.emit('error', err.message);
    }
  });

  socket.on('disconnect', () => {
    Object.keys(roomUsers).forEach((sheetId) => {
      if (roomUsers[sheetId][socket.id]) {
        delete roomUsers[sheetId][socket.id];
        io.to(sheetId).emit('room-users', Object.values(roomUsers[sheetId]));
      }
    });
  });
});

// Weekly backup cron (every Sunday at 2am)
const BACKUP_DIR = process.env.BACKUP_DIR || './backups';
mkdirSync(BACKUP_DIR, { recursive: true });

// Weekly auto-backup of ALL tables every Sunday at 2am
cron.schedule('0 2 * * 0', async () => {
  try {
    const { rows: sheets } = await query('SELECT * FROM spreadsheets');
    if (!sheets.length) return;

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `weekly-backup-${timestamp}.zip`;
    const filepath = join(BACKUP_DIR, filename);
    const output = createWriteStream(filepath);
    const archive = archiver('zip');
    archive.pipe(output);

    for (const sheet of sheets) {
      const { rows: dataRows } = await query(
        'SELECT data FROM spreadsheet_data WHERE spreadsheet_id = $1 ORDER BY sheet_index',
        [sheet.id]
      );
      const buffer = await exportExcel(dataRows.map((r) => r.data));
      archive.append(Buffer.from(buffer), { name: `${sheet.name}.xlsx` });
    }

    await archive.finalize();
    output.on('close', async () => {
      await query(
        'INSERT INTO backups (filename, size_bytes, backup_type) VALUES ($1, $2, $3)',
        [filename, archive.pointer(), 'weekly']
      );
    });
  } catch (err) {
    console.error('Auto backup failed:', err);
  }
});

const PORT = process.env.PORT || 3001;
httpServer.listen(PORT, () => console.log(`Server running on port ${PORT}`));

// ── Diff helpers ──────────────────────────────────────────────────────────────

function colNameFromIndex(c) {
  let name = '';
  c += 1;
  while (c > 0) { c -= 1; name = String.fromCharCode(65 + (c % 26)) + name; c = Math.floor(c / 26); }
  return name;
}

// Flatten sheet data (as sent by the frontend) into a { "r_c": cellObj } map.
// The frontend saves data.cells = cellsFromSheet(s) which is already flat { "r_c": {...} }.
function flattenSavedCells(data) {
  if (data && typeof data === 'object' && data.cells) return data.cells;
  return {};
}

function computeSheetDiff(prevDataRaw, currData) {
  const prevCells = flattenSavedCells(
    typeof prevDataRaw === 'string' ? JSON.parse(prevDataRaw) : prevDataRaw
  );
  const currCells = flattenSavedCells(currData);

  const allKeys = new Set([...Object.keys(prevCells), ...Object.keys(currCells)]);
  const changes = [];

  for (const key of allKeys) {
    const pv = String(prevCells[key]?.v ?? prevCells[key]?.m ?? '').trim();
    const cv = String(currCells[key]?.v ?? currCells[key]?.m ?? '').trim();
    if (pv === cv) continue;
    if (pv === '' && cv === '') continue;
    const [r, c] = key.split('_').map(Number);
    changes.push({
      key,
      col: `${colNameFromIndex(c)}${r + 1}`,
      newVal: cv.slice(0, 100),
      oldVal: pv.slice(0, 100),
    });
    if (changes.length >= 200) break;
  }

  let summary = null;
  if (changes.length > 0) {
    const examples = changes.slice(0, 3).map(ch =>
      `${ch.col}${ch.newVal ? ` ("${ch.newVal.slice(0, 20)}")` : ' (удалено)'}`
    );
    summary = `Изменено ячеек: ${changes.length} (${examples.join(', ')})`;
  }

  return { cells: changes, summary };
}
