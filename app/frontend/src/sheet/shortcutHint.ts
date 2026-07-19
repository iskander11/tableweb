/** Human-readable shortcut label, e.g. "Ctrl + Shift + Z". */
export function formatShortcutLabel(e: KeyboardEvent): string {
  const parts: string[] = [];
  if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
  if (e.shiftKey) parts.push('Shift');
  if (e.altKey) parts.push('Alt');

  let key = e.key;
  if (key === ' ') key = 'Space';
  else if (key.length === 1) key = key.toUpperCase();
  else if (key.startsWith('Arrow')) key = key.replace('Arrow', '');
  else if (key === 'Delete') key = 'Del';
  else if (key === 'Escape') key = 'Esc';

  if (!['Control', 'Shift', 'Alt', 'Meta'].includes(key)) {
    parts.push(key);
  }
  return parts.join(' + ');
}

export const FORTUNE_DIALOG_SELECTOR =
  '#fortune-search-replace, .fortune-search-replace, .fortune-dialog, .fortune-modal-container, .fortune-popover-backdrop, .tw-format-cells-dialog, .tw-format-cells-backdrop';
