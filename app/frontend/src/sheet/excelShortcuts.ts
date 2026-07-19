/** Helpers so Excel-like shortcuts reach FortuneSheet even when focus is on the page chrome. */

import { formatShortcutLabel, FORTUNE_DIALOG_SELECTOR } from './shortcutHint';

export { formatShortcutLabel };

export function ensureFortuneContainerFocusable(wrap: HTMLElement): void {
  const container = wrap.querySelector('.fortune-container') as HTMLElement | null;
  if (!container || container.getAttribute('tabindex') != null) return;
  container.setAttribute('tabindex', '-1');
  container.style.outline = 'none';
}

export function focusFortuneSheet(wrap: HTMLElement): void {
  const cellInput = wrap.querySelector('#luckysheet-rich-text-editor') as HTMLElement | null;
  const container = wrap.querySelector('.fortune-container') as HTMLElement | null;
  (cellInput || container)?.focus?.();
}

export function focusSearchDialogInput(wrap: HTMLElement): void {
  window.setTimeout(() => {
    const input = wrap.querySelector('#fortune-search-replace input') as HTMLInputElement | null;
    input?.focus();
    input?.select();
  }, 0);
}

function isInsideFortuneSheet(target: EventTarget | null, wrap: HTMLElement): boolean {
  if (!(target instanceof Node)) return false;
  return wrap.contains(target);
}

function isFortuneDialogTarget(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && !!target.closest(FORTUNE_DIALOG_SELECTOR);
}

export function isExternalFormTarget(target: EventTarget | null, wrap: HTMLElement): boolean {
  if (!(target instanceof HTMLElement)) return false;
  if (isFortuneDialogTarget(target)) return true;
  if (target.closest('.fortune-container, .fortune-workarea')) return false;
  if (target.closest('header')) return true;
  if (target.closest('[data-sheet-history]')) return true;
  return !!target.closest('input, textarea, select');
}

/** Keys FortuneSheet handles like Excel (see fortune-sheet keyboard.ts). */
export function isSpreadsheetShortcut(e: KeyboardEvent): boolean {
  const ctrl = e.ctrlKey || e.metaKey;

  if (ctrl && e.code === 'KeyS') return true;
  if (ctrl && ['KeyC', 'KeyV', 'KeyX', 'KeyZ', 'KeyY', 'KeyB', 'KeyA', 'KeyF', 'KeyH', 'KeyD', 'KeyR'].includes(e.code)) {
    return true;
  }
  if (ctrl && e.shiftKey && e.code === 'KeyZ') return true;
  if (ctrl && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) return true;

  if (['Delete', 'Backspace', 'Enter', 'Tab', 'F2', 'Escape'].includes(e.key)) return true;
  if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) return true;
  if (e.shiftKey && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) return true;

  return false;
}

export function fortuneSheetHasKeyboardFocus(wrap: HTMLElement): boolean {
  const active = document.activeElement;
  if (!active) return false;
  if (wrap.contains(active)) {
    if (active.closest('.fortune-container, .fortune-workarea')) return true;
  }
  return false;
}

export function forwardKeyboardEventToFortune(wrap: HTMLElement, e: KeyboardEvent): void {
  const container = wrap.querySelector('.fortune-container') as HTMLElement | null;
  if (!container) return;
  focusFortuneSheet(wrap);
  container.dispatchEvent(
    new KeyboardEvent('keydown', {
      key: e.key,
      code: e.code,
      ctrlKey: e.ctrlKey,
      shiftKey: e.shiftKey,
      altKey: e.altKey,
      metaKey: e.metaKey,
      bubbles: true,
      cancelable: true,
    }),
  );
}

export type ExcelShortcutOptions = {
  wrap: HTMLElement;
  enabled: boolean;
  onSave: () => void;
  onFormatCells?: () => void;
  onShortcutHint?: (label: string) => void;
};

export function bindExcelShortcuts({ wrap, enabled, onSave, onFormatCells, onShortcutHint }: ExcelShortcutOptions): () => void {
  ensureFortuneContainerFocusable(wrap);

  const onKeyDown = (e: KeyboardEvent) => {
    if (!enabled) return;
    if (isExternalFormTarget(e.target, wrap)) return;

    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.code === 'Digit1') {
      e.preventDefault();
      e.stopPropagation();
      onShortcutHint?.(formatShortcutLabel(e));
      onFormatCells?.();
      return;
    }

    if (!isSpreadsheetShortcut(e)) return;

    onShortcutHint?.(formatShortcutLabel(e));

    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.code === 'KeyS') {
      e.preventDefault();
      e.stopPropagation();
      onSave();
      return;
    }

    if (fortuneSheetHasKeyboardFocus(wrap)) {
      if ((e.ctrlKey || e.metaKey) && e.code === 'KeyF') {
        window.setTimeout(() => focusSearchDialogInput(wrap), 0);
      }
      return;
    }

    const active = document.activeElement;
    const shouldForward =
      active === document.body ||
      active === document.documentElement ||
      active?.closest('header') != null ||
      isInsideFortuneSheet(e.target, wrap);

    if (!shouldForward) return;

    e.preventDefault();
    e.stopPropagation();
    forwardKeyboardEventToFortune(wrap, e);

    if ((e.ctrlKey || e.metaKey) && e.code === 'KeyF') {
      focusSearchDialogInput(wrap);
    }
  };

  const onPointerDown = (e: Event) => {
    if (!enabled) return;
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    if (!wrap.contains(target)) return;
    if (target.closest(`header, [data-sheet-history], ${FORTUNE_DIALOG_SELECTOR}`)) return;
    if (target.closest('input, textarea, select, button')) return;
    if (target.closest('.fortune-container, .fortune-workarea, .fortune-cell-area, canvas.fortune-sheet-canvas')) {
      ensureFortuneContainerFocusable(wrap);
      focusFortuneSheet(wrap);
    }
  };

  window.addEventListener('keydown', onKeyDown, true);
  wrap.addEventListener('mousedown', onPointerDown, true);
  return () => {
    window.removeEventListener('keydown', onKeyDown, true);
    wrap.removeEventListener('mousedown', onPointerDown, true);
  };
};
