// @ts-expect-error ssf has no bundled types
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
      { fa: '#,##0.00 ₽', label: '# ##0,00 ₽' },
      { fa: '#,##0 ₽;-# ##0 ₽', label: '# ##0 ₽;-# ##0 ₽' },
      { fa: '#,##0.00" ₽"', label: '# ##0,00 "₽"' },
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

/** Convert Russian Excel format tokens to SSF/English codes. */
export function normalizeFormatCode(fa: string): string {
  let s = fa.trim();
  if (!s) return 'General';

  const replacements: [RegExp, string][] = [
    [/ГГГГ/g, 'yyyy'],
    [/ГГ/g, 'yy'],
    [/ДД/g, 'dd'],
    [/Д/g, 'd'],
    [/ММММ/g, 'mmmm'],
    [/МММ/g, 'mmm'],
    [/ММ/g, 'mm'],
    [/М/g, 'm'],
    [/ЧЧ/g, 'HH'],
    [/чч/g, 'hh'],
    [/Ч/g, 'H'],
    [/ч/g, 'h'],
    [/СС/g, 'ss'],
    [/сс/g, 'ss'],
    [/С/g, 's'],
    [/с/g, 's'],
    [/\[Красный\]/gi, '[Red]'],
    [/\[Синий\]/gi, '[Blue]'],
    [/\[Зеленый\]/gi, '[Green]'],
  ];
  for (const [re, rep] of replacements) s = s.replace(re, rep);
  return s;
}

export function inferCellType(fa: string): string {
  const code = normalizeFormatCode(fa);
  if (code === '@') return 's';
  if (code === 'General') return 'g';
  if (/[dmyhHsSb]/i.test(code) && !/^#/.test(code) && !/[%₽$€]/.test(code)) return 'd';
  return 'n';
}

function sampleValueForFormat(fa: string, raw: unknown): number | string | boolean {
  const code = normalizeFormatCode(fa);
  if (code === '@') return typeof raw === 'string' && raw ? raw : 'Текст';
  if (typeof raw === 'boolean') return raw;
  if (raw != null && raw !== '' && !Number.isNaN(Number(raw))) return Number(raw);
  if (/[dmyhH]/i.test(code)) return 45292.5125; // ~2024-01-15 12:18
  if (code.includes('%')) return 0.1234;
  return 1234.567;
}

export function formatPreview(fa: string, rawValue?: unknown): string {
  const code = normalizeFormatCode(fa);
  if (!code || code === 'General') {
    if (rawValue == null || rawValue === '') return '';
    return String(rawValue);
  }
  if (code === '@') return String(sampleValueForFormat(code, rawValue));
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
  for (const cat of FORMAT_CATEGORIES) {
    if (cat.formats.some((f) => f.fa === fa)) return cat.id;
  }
  return 'custom';
}

export function applyCellFormatToWorkbook(workbook: any, fa: string): boolean {
  if (!workbook?.getSelection || !workbook?.setCellFormat) return false;
  const selections = workbook.getSelection() || [];
  if (!selections.length) return false;

  const ssfFa = normalizeFormatCode(fa);
  const ct = { fa: ssfFa, t: inferCellType(ssfFa) };

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
  return true;
}
