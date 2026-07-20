import type { Socket } from 'socket.io-client';
import api from '../api/client';

export type SheetSavePayload = {
  sheetIndex: number;
  data: Record<string, unknown>;
  logChange: boolean;
};

const SAVE_ACK_MS = 25000;

/** Finish in-cell edit so FortuneSheet commits the value before we read sheet data. */
export function commitActiveCellEdit(wrap: HTMLElement | null): void {
  if (!wrap) return;

  const active = document.activeElement;
  const editorSelectors = [
    '#luckysheet-rich-text-editor',
    '#fortune-formula-input',
    '.fortune-formula-input',
    'div[data-sheet-editor="true"]',
  ];

  for (const sel of editorSelectors) {
    const el = wrap.querySelector(sel) as HTMLElement | null;
    if (!el) continue;
    if (active === el || el.contains(active)) {
      el.blur();
    }
  }

  if (active instanceof HTMLElement && wrap.contains(active)) {
    if (active.isContentEditable || active.tagName === 'INPUT' || active.tagName === 'TEXTAREA') {
      active.blur();
    }
  }

  // Click grid container so FortuneSheet exits edit mode (same as clicking another cell).
  const container = wrap.querySelector('.fortune-container') as HTMLElement | null;
  container?.focus();
}

export function waitForSheetCommit(ms = 120): Promise<void> {
  return new Promise((resolve) => { window.setTimeout(resolve, ms); });
}

function saveOneViaSocket(
  socket: Socket,
  sheetId: string,
  payload: SheetSavePayload,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const timer = window.setTimeout(() => {
      reject(new Error(`Таймаут сохранения (лист ${payload.sheetIndex + 1})`));
    }, SAVE_ACK_MS);

    socket.emit(
      'save-sheet',
      {
        sheetId,
        sheetIndex: payload.sheetIndex,
        data: payload.data,
        logChange: payload.logChange,
      },
      (resp: { ok?: boolean; error?: string } | undefined) => {
        window.clearTimeout(timer);
        if (resp?.ok) {
          resolve();
          return;
        }
        reject(new Error(resp?.error || 'Сервер не подтвердил сохранение'));
      },
    );
  });
}

async function saveOneViaApi(
  sheetId: string,
  payload: SheetSavePayload,
): Promise<void> {
  try {
    await api.put(`/spreadsheets/${sheetId}/sheets/${payload.sheetIndex}`, {
      data: payload.data,
      logChange: payload.logChange,
    });
  } catch (err: unknown) {
    const ax = err as { response?: { data?: { error?: string } }; message?: string };
    throw new Error(ax.response?.data?.error || ax.message || 'Ошибка сохранения');
  }
}

/** Save all sheet tabs; socket first, HTTP fallback if socket unavailable or fails. */
export async function saveSheetPayloads(
  socket: Socket | null | undefined,
  sheetId: string,
  payloads: SheetSavePayload[],
): Promise<void> {
  if (!payloads.length) {
    throw new Error('Нет данных для сохранения');
  }

  const trySocket = socket?.connected === true;

  if (trySocket && socket) {
    try {
      await Promise.all(payloads.map((p) => saveOneViaSocket(socket, sheetId, p)));
      return;
    } catch (err) {
      // Fall back to REST — e.g. ack timeout or transient socket error.
      try {
        for (const p of payloads) {
          await saveOneViaApi(sheetId, p);
        }
        return;
      } catch {
        throw err;
      }
    }
  }

  for (const p of payloads) {
    await saveOneViaApi(sheetId, p);
  }
}
