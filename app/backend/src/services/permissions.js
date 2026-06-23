import { query } from '../db/index.js';

// Single source of truth for the role model. Two distinct rights:
//
// EDIT (change cell data / import): the workbook is a shared corporate base.
//   - admin  → any table
//   - editor → any table
//   - reader → never
//
// MANAGE (delete / backup / rename): scoped to ownership.
//   - admin  → any table
//   - editor → only tables they created (their own)
//   - reader → never

// Edit rights: editors and admins can edit every table; readers never.
export function canEditSheet(user) {
  return !!user && (user.role === 'admin' || user.role === 'editor');
}

// Manage rights (delete / backup / rename): admins anywhere, editors only on
// tables they own, readers never.
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

// Who is allowed to create new tables: editors and admins (not readers).
export function canCreateSheet(user) {
  return !!user && (user.role === 'admin' || user.role === 'editor');
}
