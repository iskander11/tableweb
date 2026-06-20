import ExcelJS from 'exceljs';

// ARGB (ExcelJS format: FFRRGGBB) -> CSS hex (#RRGGBB)
const argbToCss = (argb) => {
  if (!argb || argb.length < 6) return null;
  // ExcelJS gives 8-char ARGB: FF at start = fully opaque
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  if (hex === '000000' || hex === '000') return '#000000';
  return `#${hex.toUpperCase()}`;
};

// CSS hex -> ARGB
const cssToArgb = (css) => {
  if (!css) return 'FF000000';
  const hex = css.replace('#', '').toUpperCase();
  return `FF${hex.padEnd(6, '0')}`;
};

// ExcelJS border style -> FortuneSheet border style
const borderStyle = (border) => {
  if (!border?.style) return null;
  const color = argbToCss(border.color?.argb) || '#000000';
  return { style: border.style === 'medium' ? '2' : '1', color };
};

// Get actual cell value (handling formula cells)
const getCellValue = (cell) => {
  const val = cell.value;
  if (val === null || val === undefined) return { v: null, f: null };

  // Formula cell
  if (typeof val === 'object' && val !== null && 'formula' in val) {
    const formula = val.formula || val.sharedFormula || '';
    const result = val.result;
    // Handle error results
    if (result && typeof result === 'object' && result.error) {
      return { v: result.error, f: formula };
    }
    return { v: result ?? '', f: formula };
  }

  // Date
  if (val instanceof Date) {
    return { v: val.toISOString(), f: null };
  }

  // Rich text
  if (typeof val === 'object' && 'richText' in val) {
    const text = val.richText.map((r) => r.text).join('');
    return { v: text, f: null };
  }

  return { v: val, f: null };
};

// Get cell type for FortuneSheet: n=number, s=string, b=boolean
const getCellType = (cell, v) => {
  if (v === null || v === undefined) return null;
  const t = cell.type;
  if (t === 4) return 'b'; // boolean
  if (typeof v === 'number') return 'n';
  if (typeof v === 'boolean') return 'b';
  return 's';
};

export async function importExcel(buffer, onProgress) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheets = [];
  const totalSheets = workbook.worksheets.length;
  let sheetsDone = 0;

  for (const worksheet of workbook.worksheets) {
    const cells = {};
    const columnWidths = {};
    const rowHeights = {};

    // --- Column widths ---
    // ExcelJS width unit is "characters" at default font (Calibri 11)
    // 1 character ≈ 7px. Default Excel column width = 8.43 chars ≈ 64px
    worksheet.columns.forEach((col, idx) => {
      if (col && col.width && col.width > 0) {
        columnWidths[idx] = Math.round(col.width * 7.5);
      }
    });

    const totalRows = worksheet.rowCount;
    let rowsDone = 0;

    // --- Rows, cells, styles ---
    worksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      const r = rowNum - 1;

      // Row height: ExcelJS uses points (pt), FortuneSheet uses pixels
      // 1pt ≈ 1.333px
      if (row.height && row.height > 0) {
        rowHeights[r] = Math.round(row.height * 1.333);
      }

      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const c = colNum - 1;
        const key = `${r}_${c}`;

        const { v, f } = getCellValue(cell);

        // Skip completely empty cells
        if (v === null && !f && !cell.style) return;

        const cellData = {};

        // Value
        cellData.v = v;
        cellData.m = cell.text ?? (v !== null ? String(v) : '');

        // Type
        const t = getCellType(cell, v);
        if (t) cellData.t = t;

        // Formula (FortuneSheet stores without leading '=')
        if (f) {
          cellData.f = f.startsWith('=') ? f.slice(1) : f;
        }

        // --- Styles ---
        const style = cell.style || {};

        // Background color
        const fill = style.fill || cell.fill;
        if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
          const fg = fill.fgColor;
          if (fg) {
            let color = null;
            if (fg.argb) color = argbToCss(fg.argb);
            else if (fg.theme !== undefined) {
              // Theme colors - use common defaults
              const themeColors = ['#FFFFFF','#000000','#E7E6E6','#44546A','#4472C4','#ED7D31','#A9D18E','#FF0000','#FFFF00','#00B0F0'];
              color = themeColors[fg.theme] || null;
            }
            if (color && color !== '#FFFFFF') cellData.bg = color;
          }
        }

        // Font
        const font = style.font || cell.font;
        if (font) {
          if (font.bold) cellData.bl = 1;
          if (font.italic) cellData.it = 1;
          if (font.underline) cellData.un = 1;
          if (font.size) cellData.fs = font.size;
          if (font.name) cellData.ff = font.name;
          if (font.color) {
            let fc = null;
            if (font.color.argb) fc = argbToCss(font.color.argb);
            else if (font.color.theme !== undefined) {
              const themeColors = ['#FFFFFF','#000000','#E7E6E6','#44546A','#4472C4','#ED7D31','#A9D18E','#FF0000','#FFFF00','#00B0F0'];
              fc = themeColors[font.color.theme] || null;
            }
            if (fc && fc !== '#000000') cellData.fc = fc;
          }
        }

        // Alignment
        const alignment = style.alignment || cell.alignment;
        if (alignment) {
          if (alignment.horizontal) {
            cellData.ht = { left: 1, center: 0, right: 2 }[alignment.horizontal] ?? 1;
          }
          if (alignment.vertical) {
            cellData.vt = { middle: 0, top: 1, bottom: 2 }[alignment.vertical] ?? 0;
          }
          if (alignment.wrapText) cellData.tb = 2;
        }

        // Borders
        const border = style.border || cell.border;
        if (border) {
          const bd = {};
          if (border.top) bd.t = borderStyle(border.top);
          if (border.bottom) bd.b = borderStyle(border.bottom);
          if (border.left) bd.l = borderStyle(border.left);
          if (border.right) bd.r = borderStyle(border.right);
          if (Object.keys(bd).length) cellData.bd = bd;
        }

        // Number format
        const numFmt = style.numFmt || cell.numFmt;
        if (numFmt) cellData.fm = numFmt;

        cells[key] = cellData;
      });

      rowsDone++;
      if (onProgress && totalRows > 0) {
        const pct = Math.round(((sheetsDone / totalSheets) + (rowsDone / totalRows / totalSheets)) * 100);
        onProgress(pct);
      }
    });

    // --- Merged cells ---
    const merges = {};
    const mergeList = worksheet.model?.merges || [];
    mergeList.forEach((mergeStr) => {
      try {
        const [startRef, endRef] = mergeStr.split(':');
        const startCell = worksheet.getCell(startRef);
        const endCell = worksheet.getCell(endRef);
        const r = startCell.row - 1;
        const c = startCell.col - 1;
        const rs = endCell.row - startCell.row + 1;
        const cs = endCell.col - startCell.col + 1;
        if (rs > 1 || cs > 1) {
          merges[`${r}_${c}`] = { r, c, rs, cs };
        }
      } catch {
        // skip invalid merge refs
      }
    });

    sheets.push({
      name: worksheet.name,
      cells,
      columnWidths,
      rowHeights,
      merges,
    });

    sheetsDone++;
    if (onProgress) onProgress(Math.round((sheetsDone / totalSheets) * 100));
  }

  return sheets;
}

export async function exportExcel(sheetsData) {
  const workbook = new ExcelJS.Workbook();

  for (const sheet of sheetsData) {
    const worksheet = workbook.addWorksheet(sheet.name || 'Sheet1');

    // Column widths (px -> chars)
    if (sheet.columnWidths) {
      Object.entries(sheet.columnWidths).forEach(([idx, px]) => {
        const colNum = parseInt(idx) + 1;
        worksheet.getColumn(colNum).width = Math.round(px / 7.5 * 10) / 10;
      });
    }

    // Row heights (px -> pts)
    if (sheet.rowHeights) {
      Object.entries(sheet.rowHeights).forEach(([rowIdx, px]) => {
        worksheet.getRow(parseInt(rowIdx) + 1).height = Math.round(px / 1.333 * 10) / 10;
      });
    }

    // Cells
    if (sheet.cells) {
      Object.entries(sheet.cells).forEach(([key, cd]) => {
        const [r, c] = key.split('_').map(Number);
        const cell = worksheet.getCell(r + 1, c + 1);

        // Value / formula
        if (cd.f) {
          cell.value = { formula: cd.f, result: cd.v };
        } else if (cd.v !== undefined && cd.v !== null) {
          cell.value = cd.v;
        }

        // Font
        const fontStyle = {};
        if (cd.bl) fontStyle.bold = true;
        if (cd.it) fontStyle.italic = true;
        if (cd.un) fontStyle.underline = true;
        if (cd.fs) fontStyle.size = cd.fs;
        if (cd.ff) fontStyle.name = cd.ff;
        if (cd.fc) fontStyle.color = { argb: cssToArgb(cd.fc) };
        if (Object.keys(fontStyle).length) cell.font = fontStyle;

        // Background
        if (cd.bg) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cssToArgb(cd.bg) } };
        }

        // Alignment
        const alignStyle = {};
        if (cd.ht !== undefined) alignStyle.horizontal = ['center', 'left', 'right'][cd.ht] || 'left';
        if (cd.vt !== undefined) alignStyle.vertical = ['middle', 'top', 'bottom'][cd.vt] || 'middle';
        if (cd.tb === 2) alignStyle.wrapText = true;
        if (Object.keys(alignStyle).length) cell.alignment = alignStyle;

        // Borders
        if (cd.bd) {
          const border = {};
          const toBorder = (b) => b ? { style: b.style === '2' ? 'medium' : 'thin', color: { argb: cssToArgb(b.color) } } : undefined;
          if (cd.bd.t) border.top = toBorder(cd.bd.t);
          if (cd.bd.b) border.bottom = toBorder(cd.bd.b);
          if (cd.bd.l) border.left = toBorder(cd.bd.l);
          if (cd.bd.r) border.right = toBorder(cd.bd.r);
          cell.border = border;
        }

        // Number format
        if (cd.fm) cell.numFmt = cd.fm;
      });
    }

    // Merges
    if (sheet.merges) {
      Object.values(sheet.merges).forEach(({ r, c, rs, cs }) => {
        if (rs > 1 || cs > 1) {
          try {
            worksheet.mergeCells(r + 1, c + 1, r + rs, c + cs);
          } catch {
            // skip conflicting merges
          }
        }
      });
    }
  }

  return workbook.xlsx.writeBuffer();
}
