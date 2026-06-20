import { Readable } from 'stream';

const EXCEL_SERVICE_URL = process.env.EXCEL_SERVICE_URL || 'http://localhost:5001';

/**
 * Parse Excel file via Python/openpyxl microservice.
 * Returns array of sheet objects in FortuneSheet format.
 */
export async function importExcel(buffer, onProgress) {
  onProgress?.(10);

  const formData = new FormData();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  formData.append('file', blob, 'upload.xlsx');

  onProgress?.(20);

  const response = await fetch(`${EXCEL_SERVICE_URL}/parse`, {
    method: 'POST',
    body: formData,
  });

  onProgress?.(85);

  if (!response.ok) {
    const err = await response.json().catch(() => ({ error: 'Python service error' }));
    throw new Error(err.error || 'Excel parse failed');
  }

  const data = await response.json();
  onProgress?.(95);
  return data.sheets;
}

/**
 * Export FortuneSheet data to Excel via Python/openpyxl microservice.
 * Returns ArrayBuffer.
 */
export async function exportExcel(sheetsData) {
  const response = await fetch(`${EXCEL_SERVICE_URL}/export`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sheets: sheetsData }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({ error: 'Python service error' }));
    throw new Error(err.error || 'Excel export failed');
  }

  return Buffer.from(await response.arrayBuffer());
}
