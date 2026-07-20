import { query } from '../db/index.js';

const CHANGELOG_LIMIT = Number(process.env.CHANGELOG_LIMIT) || 200;

function colNameFromIndex(c) {
  let name = '';
  c += 1;
  while (c > 0) {
    c -= 1;
    name = String.fromCharCode(65 + (c % 26)) + name;
    c = Math.floor(c / 26);
  }
  return name;
}

function flattenSavedCells(data) {
  if (data && typeof data === 'object' && data.cells) return data.cells;
  return {};
}

export function computeSheetDiff(prevDataRaw, currData) {
  const prevCells = flattenSavedCells(
    typeof prevDataRaw === 'string' ? JSON.parse(prevDataRaw) : prevDataRaw,
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
    const examples = changes.slice(0, 3).map((ch) =>
      `${ch.col}${ch.newVal ? ` ("${ch.newVal.slice(0, 20)}")` : ' (удалено)'}`,
    );
    summary = `Изменено ячеек: ${changes.length} (${examples.join(', ')})`;
  }

  return { cells: changes, summary };
}

/** Persist one worksheet tab. Returns changelog entry when logChange is true. */
export async function persistSpreadsheetSheet({
  sheetId,
  sheetIndex,
  data,
  user,
  logChange = false,
}) {
  let changedCells = null;
  let summary = null;

  if (logChange) {
    try {
      const prev = await query(
        'SELECT data FROM spreadsheet_data WHERE spreadsheet_id = $1 AND sheet_index = $2',
        [sheetId, sheetIndex],
      );
      if (prev.rows.length > 0) {
        const diff = computeSheetDiff(prev.rows[0].data, data);
        changedCells = diff.cells.length > 0 ? diff.cells : null;
        summary = diff.summary;
      }
    } catch {
      /* diff failure must not block the save */
    }
  }

  await query(
    `INSERT INTO spreadsheet_data (spreadsheet_id, sheet_index, data)
     VALUES ($1, $2, $3)
     ON CONFLICT (spreadsheet_id, sheet_index) DO UPDATE SET data = $3, updated_at = NOW()`,
    [sheetId, sheetIndex, JSON.stringify(data)],
  );
  await query('UPDATE spreadsheets SET updated_at = NOW() WHERE id = $1', [sheetId]);

  let entry = null;
  if (logChange) {
    await query(
      `INSERT INTO change_log (spreadsheet_id, sheet_index, user_id, username, summary, changed_cells)
       VALUES ($1, $2, $3, $4, $5, $6)`,
      [sheetId, sheetIndex, user.id, user.username, summary || null,
        changedCells ? JSON.stringify(changedCells) : null],
    );
    await query(
      `DELETE FROM change_log
       WHERE spreadsheet_id = $1
         AND id NOT IN (
           SELECT id FROM change_log WHERE spreadsheet_id = $1
           ORDER BY saved_at DESC LIMIT $2
         )`,
      [sheetId, CHANGELOG_LIMIT],
    ).catch(() => { /* retention is best-effort */ });

    entry = {
      username: user.username,
      sheet_index: sheetIndex,
      summary: summary || null,
      changed_cells: changedCells || null,
      saved_at: new Date().toISOString(),
    };
  }

  return { ok: true, sheetIndex, entry };
}
