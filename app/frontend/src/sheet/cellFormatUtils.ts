import SSF from 'ssf';

export type FormatCategoryId =
  | 'general'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'date'
  | 'time'
  | 'datetime'
  | 'percentage'
  | 'fraction'
  | 'scientific'
  | 'text'
  | 'custom';

export type FormatPreset = { fa: string; label: string };

export const FORMAT_CATEGORIES: { id: FormatCategoryId; label: string; formats: FormatPreset[] }[] = [
  {
    id: 'general',
    label: 'Общий',
    formats: [{ fa: 'General', label: 'Основной' }],
  },
  {
    id: 'number',
    label: 'Числовой',
    formats: [
      { fa: '0', label: '0' },
      { fa: '0.00', label: '0,00' },
      { fa: '#,##0', label: '# ##0' },
      { fa: '#,##0.00', label: '# ##0,00' },
      { fa: '#,##0;[Red]-#,##0', label: '# ##0;[Красный]-# ##0' },
      { fa: '#,##0.00;[Red]-#,##0.00', label: '# ##0,00;[Красный]-# ##0,00' },
    ],
  },
  {
    id: 'currency',
    label: 'Денежный',
    formats: [
      { fa: '#,##0.00" ₽"', label: '# ##0,00 ₽' },
      { fa: '#,##0" ₽";-#,##0" ₽"', label: '# ##0 ₽;-# ##0 ₽' },
    ],
  },
  {
    id: 'accounting',
    label: 'Финансовый',
    formats: [
      { fa: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', label: 'Финансовый' },
    ],
  },
  {
    id: 'date',
    label: 'Дата',
    formats: [
      { fa: 'dd.mm.yyyy', label: 'ДД.ММ.ГГГГ' },
      { fa: 'd.m.yyyy', label: 'Д.М.ГГГГ' },
      { fa: 'yyyy-mm-dd', label: 'ГГГГ-ММ-ДД' },
      { fa: 'dd.mm.yy', label: 'ДД.ММ.ГГ' },
    ],
  },
  {
    id: 'time',
    label: 'Время',
    formats: [
      { fa: 'h:mm', label: 'ч:мм' },
      { fa: 'h:mm:ss', label: 'ч:мм:сс' },
      { fa: 'hh:mm', label: 'чч:мм' },
      { fa: 'hh:mm:ss', label: 'чч:мм:сс' },
    ],
  },
  {
    id: 'datetime',
    label: 'Дата и время',
    formats: [
      { fa: 'dd.mm.yyyy h:mm', label: 'ДД.ММ.ГГГГ ч:мм' },
      { fa: 'dd.mm.yyyy hh:mm', label: 'ДД.ММ.ГГГГ чч:мм' },
      { fa: 'dd.mm.yy hh:mm', label: 'ДД.ММ.ГГ чч:мм' },
      { fa: 'yyyy-mm-dd hh:mm', label: 'ГГГГ-ММ-ДД чч:мм' },
    ],
  },
  {
    id: 'percentage',
    label: 'Процентный',
    formats: [
      { fa: '0%', label: '0%' },
      { fa: '0.00%', label: '0,00%' },
    ],
  },
  {
    id: 'fraction',
    label: 'Дробный',
    formats: [
      { fa: '# ?/?', label: '# ?/?' },
      { fa: '# ??/??', label: '# ??/??' },
    ],
  },
  {
    id: 'scientific',
    label: 'Экспоненциальный',
    formats: [{ fa: '0.00E+00', label: '0,00E+00' }],
  },
  {
    id: 'text',
    label: 'Текстовый',
    formats: [{ fa: '@', label: '@' }],
  },
  {
    id: 'custom',
    label: '(все форматы)',
    formats: [],
  },
];

const BUILTIN_FA_LIST: string[] = FORMAT_CATEGORIES.flatMap((c) => c.formats.map((f) => f.fa));

const CUSTOM_FA_STORAGE_KEY = 'tableweb_custom_cell_formats';

/** Excel max serial date ~ 9999-12-31 */
const MAX_EXCEL_DATE_SERIAL = 2958465;

/** Known invalid presets saved before the fix — map to SSF-safe codes. */
const LEGACY_FORMAT_FIXES: Record<string, string> = {
  '#,##0.00 ₽': '#,##0.00" ₽"',
  '#,##0 ₽;-# ##0 ₽': '#,##0" ₽";-#,##0" ₽"',
};

export type FormatApplyResult = { ok: true } | { ok: false; error: string };

/** Convert Russian Excel format tokens to SSF/English codes (upper & lower case). */
export function normalizeFormatCode(fa: string): string {
  let s = fa.trim();
  if (!s) return 'General';

  const replacements: [RegExp, string][] = [
    [/ГГГГ|гггг/g, 'yyyy'],
    [/ГГ|гг/g, 'yy'],
    [/ДД|дд/g, 'dd'],
    [/ММММ|мммм/g, 'mmmm'],
    [/МММ|ммм/g, 'mmm'],
    [/ММ|мм/g, 'mm'],
    [/ЧЧ|чч/g, 'hh'],
    [/ХХ|хх/g, 'hh'],
    [/Ч|ч/g, 'h'],
    [/Х|х/g, 'h'],
    [/Д|д/g, 'd'],
    [/М|м/g, 'm'],
    [/СС|сс/g, 'ss'],
    [/С|с/g, 's'],
    [/\[Красный\]/gi, '[Red]'],
    [/\[Синий\]/gi, '[Blue]'],
    [/\[Зеленый\]/gi, '[Green]'],
  ];
  for (const [re, rep] of replacements) s = s.replace(re, rep);
  return s;
}

/** Make a format string safe for FortuneSheet/SSF (quote bare currency symbols, etc.). */
export function sanitizeFormatCode(fa: string): string {
  let code = normalizeFormatCode(fa);
  if (LEGACY_FORMAT_FIXES[code]) return LEGACY_FORMAT_FIXES[code];

  code = code.replace(/(\d|\?|0|#|"|%)([ \u00A0])([₽$€])(?=($|;|\)|,))/g, '$1$2"$3"');
  code = code.replace(/-#\s+##0\s*([₽$€])/g, '-#,##0"$1"');

  return code;
}

export function safeFormatValue(fa: string, v: unknown): string | null {
  if (v == null || v === '') return '';
  const code = sanitizeFormatCode(fa);
  try {
    return SSF.format(code, v as string | number);
  } catch {
    return null;
  }
}

export function isValidFormatCode(fa: string): boolean {
  const code = sanitizeFormatCode(fa);
  if (!code) return false;
  if (code === 'General' || code === '@') return true;
  try {
    SSF.format(code, 0);
    SSF.format(code, 1234.567);
    SSF.format(code, 45292.5125);
    return true;
  } catch {
    return false;
  }
}

export function isDateFormatCode(fa: string): boolean {
  return inferCellType(fa) === 'd';
}

function parseAsExcelSerial(v: unknown): number | null {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (typeof v === 'string') {
    const t = v.trim();
    if (!t) return null;
    if (/^\d+$/.test(t) && t.length > 9) return null;
    const n = Number(t.replace(/\s/g, '').replace(',', '.'));
    if (!Number.isFinite(n)) return null;
    return n;
  }
  return null;
}

/** Date/time formats only accept Excel serial dates (roughly 1900–9999). */
export function valueCompatibleWithFormat(fa: string, v: unknown): boolean {
  const code = sanitizeFormatCode(fa);
  if (code === '@' || code === 'General') return true;
  if (v == null || v === '') return true;
  if (!isValidFormatCode(code)) return false;

  if (isDateFormatCode(code)) {
    const serial = parseAsExcelSerial(v);
    if (serial == null) return false;
    if (serial < 0 || serial > MAX_EXCEL_DATE_SERIAL) return false;
    return safeFormatValue(code, serial) != null;
  }

  return safeFormatValue(code, v) != null;
}

export function inferCellType(fa: string): string {
  const code = sanitizeFormatCode(fa);
  if (code === '@') return 's';
  if (code === 'General') return 'g';
  try {
    if (SSF.is_date(code)) return 'd';
  } catch {
    // fall through
  }
  if (/[dmyhH]/i.test(code) && !/^#/.test(code) && !/[%₽$€]/.test(code)) return 'd';
  return 'n';
}

/** Fix persisted cell ct so FortuneSheet won't crash on edit (SSF.format throws on bad fa/v). */
export function sanitizeCellFormatValue(cell: any): any {
  if (!cell || typeof cell !== 'object' || !cell.ct?.fa) return cell;
  const fa = sanitizeFormatCode(String(cell.ct.fa));
  if (!isValidFormatCode(fa)) {
    const m = cell.v != null && cell.v !== '' ? String(cell.v) : cell.m;
    return { ...cell, ct: { fa: 'General', t: 'g' }, ...(m != null ? { m } : {}) };
  }

  if (cell.v != null && cell.v !== '' && !valueCompatibleWithFormat(fa, cell.v)) {
    return { ...cell, ct: { fa: '@', t: 's' }, m: String(cell.v) };
  }

  const t = inferCellType(fa);
  const m = cell.v != null && cell.v !== ''
    ? (safeFormatValue(fa, cell.v) ?? String(cell.v))
    : cell.m;

  if (cell.ct.fa === fa && cell.ct.t === t && (m == null || cell.m === m)) return cell;
  return { ...cell, ct: { ...cell.ct, fa, t }, ...(m != null ? { m } : {}) };
}

function sampleValueForFormat(fa: string, raw: unknown): number | string | boolean {
  const code = sanitizeFormatCode(fa);
  if (code === '@') return typeof raw === 'string' && raw ? raw : 'Текст';
  if (typeof raw === 'boolean') return raw;
  if (raw != null && raw !== '' && !Number.isNaN(Number(raw))) {
    const n = Number(raw);
    if (isDateFormatCode(code) && (n < 0 || n > MAX_EXCEL_DATE_SERIAL)) return 45292.5125;
    return n;
  }
  if (/[dmyhH]/i.test(code)) return 45292.5125;
  if (code.includes('%')) return 0.1234;
  return 1234.567;
}

export function formatPreview(fa: string, rawValue?: unknown): string {
  const code = sanitizeFormatCode(fa);
  if (!code || code === 'General') {
    if (rawValue == null || rawValue === '') return '';
    return String(rawValue);
  }
  if (code === '@') return String(sampleValueForFormat(code, rawValue));
  if (rawValue != null && rawValue !== '' && !valueCompatibleWithFormat(code, rawValue)) {
    return `${rawValue} (не подходит для формата)`;
  }
  try {
    const v = sampleValueForFormat(code, rawValue);
    return SSF.format(code, v);
  } catch {
    return '—';
  }
}

export function loadCustomFormats(): string[] {
  try {
    const raw = localStorage.getItem(CUSTOM_FA_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter((x) => typeof x === 'string') : [];
  } catch {
    return [];
  }
}

export function saveCustomFormat(fa: string): void {
  const code = fa.trim();
  if (!code || code === 'General' || BUILTIN_FA_LIST.includes(code)) return;
  const list = loadCustomFormats().filter((x) => x !== code);
  list.unshift(code);
  localStorage.setItem(CUSTOM_FA_STORAGE_KEY, JSON.stringify(list.slice(0, 40)));
}

export function deleteCustomFormat(fa: string): void {
  const list = loadCustomFormats().filter((x) => x !== fa);
  localStorage.setItem(CUSTOM_FA_STORAGE_KEY, JSON.stringify(list));
}

export function allFormatPresets(): FormatPreset[] {
  const seen = new Set<string>();
  const out: FormatPreset[] = [];
  for (const cat of FORMAT_CATEGORIES) {
    for (const f of cat.formats) {
      if (seen.has(f.fa)) continue;
      seen.add(f.fa);
      out.push(f);
    }
  }
  for (const fa of loadCustomFormats()) {
    if (seen.has(fa)) continue;
    seen.add(fa);
    out.push({ fa, label: fa });
  }
  return out;
}

export function categoryForFormat(fa: string): FormatCategoryId {
  const safe = sanitizeFormatCode(fa);
  for (const cat of FORMAT_CATEGORIES) {
    if (cat.formats.some((f) => f.fa === fa || f.fa === safe)) return cat.id;
  }
  return 'custom';
}

function formatApplyError(fa: string, v: unknown): string {
  if (isDateFormatCode(fa)) {
    return `Значение «${v}» не является датой Excel. Формат даты/времени работает с датами (например 15.01.2024) или с серийным номером даты (от 1 до ${MAX_EXCEL_DATE_SERIAL}). Для длинных чисел используйте текстовый формат (@).`;
  }
  return `Значение «${v}» не подходит для выбранного формата.`;
}

export function applyCellFormatToWorkbook(workbook: any, fa: string): FormatApplyResult {
  if (!workbook?.getSelection || !workbook?.setCellFormat) {
    return { ok: false, error: 'Таблица не готова. Попробуйте ещё раз.' };
  }
  const selections = workbook.getSelection() || [];
  if (!selections.length) {
    return { ok: false, error: 'Сначала выделите одну или несколько ячеек на листе.' };
  }

  const ssfFa = sanitizeFormatCode(fa);
  if (!isValidFormatCode(ssfFa)) {
    return { ok: false, error: 'Некорректный код формата. Для часов используйте «ч» или «чч», не «х».' };
  }

  const sheet = workbook.getSheet?.();
  const data = sheet?.data;

  for (const sel of selections) {
    const r0 = sel.row[0];
    const r1 = sel.row[1];
    const c0 = sel.column[0];
    const c1 = sel.column[1];
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        const cell = data?.[r]?.[c];
        const v = cell && typeof cell === 'object' ? cell.v : cell;
        if (v != null && v !== '' && !valueCompatibleWithFormat(ssfFa, v)) {
          return { ok: false, error: formatApplyError(ssfFa, v) };
        }
      }
    }
  }

  const ct = { fa: ssfFa, t: inferCellType(ssfFa) };

  try {
    for (const sel of selections) {
      const range = { row: sel.row, column: sel.column };
      if (typeof workbook.setCellFormatByRange === 'function') {
        workbook.setCellFormatByRange('ct', ct, range);
        continue;
      }
      const r0 = sel.row[0];
      const r1 = sel.row[1];
      const c0 = sel.column[0];
      const c1 = sel.column[1];
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          workbook.setCellFormat(r, c, 'ct', ct);
        }
      }
    }
  } catch {
    return { ok: false, error: 'Не удалось применить формат.' };
  }
  return { ok: true };
}

/** Prevent FortuneSheet crash when editing a cell whose format cannot display the new value. */
export function guardCellBeforeUpdate(workbook: any, r: number, c: number, value: unknown): void {
  const sheet = workbook?.getSheet?.();
  const cell = sheet?.data?.[r]?.[c];
  if (!cell || typeof cell !== 'object' || !cell.ct?.fa) return;
  if (value == null || value === '') return;

  const fa = sanitizeFormatCode(String(cell.ct.fa));
  if (fa === 'General' || fa === '@') return;
  if (valueCompatibleWithFormat(fa, value)) return;

  cell.ct = { fa: '@', t: 's' };
}
