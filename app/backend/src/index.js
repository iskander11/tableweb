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

  socket.on('save-sheet', async ({ sheetId, sheetIndex, data, summary }) => {
    try {
      await query(
        `INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data)
         VALUES ($1, $2, $3)
         ON CONFLICT (spreadsheet_id, sheet_index) DO UPDATE SET data = $3, updated_at = NOW()`,
        [sheetId, sheetIndex, JSON.stringify(data)]
      );
      await query('UPDATE spreadsheets SET updated_at = NOW() WHERE id = $1', [sheetId]);
      await query(
        `INSERT INTO change_log (spreadsheet_id, sheet_index, user_id, username, summary)
         VALUES ($1, $2, $3, $4, $5)`,
        [sheetId, sheetIndex, socket.user.id, socket.user.username, summary || null]
      );
      io.to(sheetId).emit('changelog-update', {
        username: socket.user.username,
        sheet_index: sheetIndex,
        summary: summary || null,
        saved_at: new Date().toISOString(),
      });
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

cron.schedule('0 2 * * 0', async () => {
  try {
    const { rows: sheets } = await query(
      'SELECT * FROM spreadsheets WHERE backup_enabled = TRUE'
    );
    if (!sheets.length) return;

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `auto-backup-${timestamp}.zip`;
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
        'INSERT INTO backups (filename, size_bytes) VALUES ($1, $2)',
        [filename, archive.pointer()]
      );
    });
  } catch (err) {
    console.error('Auto backup failed:', err);
  }
});

const PORT = process.env.PORT || 3001;
httpServer.listen(PORT, () => console.log(`Server running on port ${PORT}`));
