// Central font registry for the spreadsheet.
//
// Two responsibilities:
//   1. Load admin-uploaded fonts into the browser (FontFace API) so the
//      FortuneSheet canvas can actually render them.
//   2. Extend FortuneSheet's font-name dropdown. The dropdown is driven by the
//      locale's `fontarray` singleton; when a user picks an item it stores the
//      *string* name as the cell's `ff`, which renders directly via CSS
//      font-family. So adding names here makes them both pickable and rendered.

import { locale } from '@fortune-sheet/core';

export interface FontMeta {
  id: string;
  displayName: string;
  familyName: string;
  format: string; // truetype | opentype | woff | woff2
  url: string;
}

// Keep FortuneSheet's original first 4 entries at the same indices so any
// legacy numeric `ff` values stored before the string migration still resolve.
const DEFAULT_FONTARRAY = ['Times New Roman', 'Arial', 'Tahoma', 'Verdana'];

// Common fonts that ship with Windows/Office — offered in the picker even
// without an upload (they render if the client has them installed).
const SYSTEM_FONTS = [
  'Calibri', 'Cambria', 'Georgia', 'Courier New', 'Comic Sans MS',
  'Trebuchet MS', 'Segoe UI', 'Garamond', 'Book Antiqua', 'Palatino Linotype',
  'Consolas', 'Impact', 'Century Gothic', 'Franklin Gothic Medium',
  'Lucida Console', 'Microsoft YaHei', 'Arial Narrow', 'Symbol', 'Wingdings',
];

const loadedFaces = new Set<string>();

function dedupePreserveOrder(names: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const n of names) {
    const key = n.trim().toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(n.trim());
  }
  return out;
}

async function loadFontFace(font: FontMeta): Promise<void> {
  if (!font.familyName || loadedFaces.has(font.familyName.toLowerCase())) return;
  if (typeof FontFace === 'undefined' || !document.fonts) return;
  try {
    const face = new FontFace(font.familyName, `url("${font.url}") format("${font.format}")`);
    await face.load();
    document.fonts.add(face);
    loadedFaces.add(font.familyName.toLowerCase());
  } catch {
    // A single bad/unsupported font file should not break the rest.
  }
}

// Push the combined font list into FortuneSheet's locale singletons so the
// toolbar dropdown shows them. Guarded: if the internal API ever changes,
// rendering (string `ff`) still works — only the picker list is affected.
function applyToFortuneSheetDropdown(names: string[]): void {
  const fontjson: Record<string, number> = {};
  names.forEach((n, i) => { fontjson[n.toLowerCase()] = i; });

  for (const lang of ['en', 'ru', 'zh']) {
    try {
      const loc: any = (locale as any)({ lang });
      if (!loc) continue;
      if (Array.isArray(loc.fontarray)) {
        loc.fontarray.splice(0, loc.fontarray.length, ...names);
      } else {
        loc.fontarray = [...names];
      }
      loc.fontjson = { ...fontjson };
    } catch {
      // ignore — dropdown extension is best-effort
    }
  }
}

/**
 * Register all fonts: load uploaded ones into the browser and refresh the
 * FortuneSheet font picker. Safe to call repeatedly.
 */
export async function registerFonts(uploaded: FontMeta[]): Promise<void> {
  await Promise.all(uploaded.map(loadFontFace));
  const names = dedupePreserveOrder([
    ...DEFAULT_FONTARRAY,
    ...SYSTEM_FONTS,
    ...uploaded.map((f) => f.familyName),
  ]);
  applyToFortuneSheetDropdown(names);
}
