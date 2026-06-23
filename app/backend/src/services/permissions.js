import { query } from '../db/index.js';

// Single source of truth for the role model:
//   - admin  → full rights on every table (edit / rename / delete / backup)
//   - editor → full rights ONLY on tables they created (their own tables)
//   - reader → no rights at all: can list and open tables read-only, nothing else
//
// Editing and managing (delete/backup/rename) follow the same rule here, so both
// resolve through canManageSheet. Per-table grants are intentionally ignored —
// the model is purely role + ownership.
export async function canManageSheet(user, spreadsheetId) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  if (user.role !== 'editor') return false; // readers (and any unknown role) never manage

  const { rows: [sheet] } = await query(
    'SELECT created_by FROM spreadsheets WHERE id = $1',
    [spreadsheetId]
  );
  if (!sheet) return false;
  return sheet.created_by === user.id;
}

// Edit rights are identical to manage rights under this model.
export const canEditSheet = canManageSheet;

// Who is allowed to create new tables: editors and admins (not readers).
export function canCreateSheet(user) {
  return !!user && (user.role === 'admin' || user.role === 'editor');
}
