import { Router } from 'express';
import multer from 'multer';
import { join, extname } from 'path';
import { mkdirSync, existsSync, unlinkSync, readFileSync } from 'fs';
import { randomUUID } from 'crypto';
import * as fontkit from 'fontkit';
import { query } from '../db/index.js';
import { authenticate, requireAdmin } from '../middleware/auth.js';

// Read the font's internal family name from the file so it matches the name
// Excel uses (e.g. "Times New Roman") without the admin having to type it.
function detectFamilyName(filePath) {
  try {
    const buf = readFileSync(filePath);
    const parsed = fontkit.create(buf);
    const font = parsed && parsed.fonts ? parsed.fonts[0] : parsed; // handle font collections
    return (font?.familyName || '').trim();
  } catch {
    return '';
  }
}

// Where uploaded font files live on disk. Served statically (no auth) so that
// the browser @font-face / FontFace loader can fetch them.
const FONTS_DIR = process.env.FONTS_DIR || './uploads/fonts';
mkdirSync(FONTS_DIR, { recursive: true });

const ALLOWED_EXT = new Set(['.ttf', '.otf', '.woff', '.woff2']);
const EXT_FORMAT = { '.ttf': 'truetype', '.otf': 'opentype', '.woff': 'woff', '.woff2': 'woff2' };

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, FONTS_DIR),
  filename: (_req, file, cb) => cb(null, `${randomUUID()}${extname(file.originalname).toLowerCase()}`),
});

const upload = multer({
  storage,
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = extname(file.originalname).toLowerCase();
    if (!ALLOWED_EXT.has(ext)) return cb(new Error('Поддерживаются только .ttf, .otf, .woff, .woff2'));
    cb(null, true);
  },
});

// Idempotent table creation (schema.sql is applied manually at deploy time).
export async function ensureFontsTable() {
  await query(`
    CREATE TABLE IF NOT EXISTS fonts (
      id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
      display_name VARCHAR(120) NOT NULL,
      family_name  VARCHAR(120) NOT NULL,
      filename     VARCHAR(200) NOT NULL,
      format       VARCHAR(20)  NOT NULL,
      uploaded_by  UUID REFERENCES users(id),
      created_at   TIMESTAMPTZ DEFAULT NOW()
    )
  `);
}

const toDto = (r) => ({
  id: r.id,
  displayName: r.display_name,
  familyName: r.family_name,
  format: r.format,
  url: `/api/fonts/files/${r.filename}`,
});

const router = Router();

// List fonts — available to every authenticated user (readers need them to
// render imported tables faithfully, not just admins).
router.get('/', authenticate, async (_req, res) => {
  try {
    const { rows } = await query(
      'SELECT id, display_name, family_name, filename, format FROM fonts ORDER BY display_name'
    );
    res.json(rows.map(toDto));
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Upload a font (admin only)
router.post('/', authenticate, requireAdmin, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Файл не передан' });
    const ext = extname(req.file.filename).toLowerCase();
    const filePath = join(FONTS_DIR, req.file.filename);

    // Auto-detect from the file; an explicit name from the admin overrides it.
    const detected = detectFamilyName(filePath);
    const familyName = (req.body.familyName || '').trim() || detected;
    const displayName = (req.body.displayName || '').trim() || familyName;

    if (!familyName) {
      if (existsSync(filePath)) { try { unlinkSync(filePath); } catch { /* ignore */ } }
      return res.status(400).json({ error: 'Не удалось определить имя шрифта — укажите его вручную' });
    }
    const { rows } = await query(
      `INSERT INTO fonts (display_name, family_name, filename, format, uploaded_by)
       VALUES ($1, $2, $3, $4, $5)
       RETURNING id, display_name, family_name, filename, format`,
      [displayName, familyName, req.file.filename, EXT_FORMAT[ext], req.user.id]
    );
    res.status(201).json(toDto(rows[0]));
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// Delete a font (admin only)
router.delete('/:id', authenticate, requireAdmin, async (req, res) => {
  try {
    const { rows } = await query('SELECT filename FROM fonts WHERE id = $1', [req.params.id]);
    if (rows.length) {
      const p = join(FONTS_DIR, rows[0].filename);
      if (existsSync(p)) { try { unlinkSync(p); } catch { /* ignore */ } }
    }
    await query('DELETE FROM fonts WHERE id = $1', [req.params.id]);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default router;
export { FONTS_DIR };
