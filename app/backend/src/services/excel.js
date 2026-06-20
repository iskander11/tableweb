import ExcelJS from 'exceljs';

// ARGB (FFRRGGBB) -> CSS hex (#RRGGBB)
const argbToCss = (argb) => {
  if (!argb || argb.length < 6) return null;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  if (/^0+$/.test(hex)) return null;
  return `#${hex.toUpperCase()}`;
};

const cssToArgb = (css) => {
  if (!css) return 'FF000000';
  return `FF${css.replace('#', '').toUpperCase().padEnd(6, '0')}`;
};

const borderStyle = (border) => {
  if (!border?.style) return null;
  const color = argbToCss(border.color?.argb) || '#000000';
  return { style: border.style === 'medium' ? '2' : '1', color };
};

// Excel date serial (days since 1900-01-01) from JS Date
const toExcelSerial = (date) => {
  const msPerDay = 86400000;
  // Days from 1900-01-01 to 1970-01-01 = 25569, +1 for Excel's leap year bug
  return Math.round(date.getTime() / msPerDay) + 25569;
};

// Format JS Date using Excel number format string
const formatDate = (date, numFmt) => {
  if (!date || !(date instanceof Date)) return '';
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = String(date.getFullYear());
  const yy = yyyy.slice(-2);
  const hh = String(date.getHours()).padStart(2, '0');
  const min = String(date.getMinutes()).padStart(2, '0');

  if (!numFmt) return `${dd}.${mm}.${yy}`;

  // Excel numFmt can have multiple sections separated by ";" (positive;negative;zero;text)
  // Take only the first section (for positive numbers / dates)
  const fmt = numFmt.split(';')[0]
    .replace(/\[.*?\]/g, '')   // remove locale brackets like [$-419]
    .replace(/\\./, '')         // remove escaped chars
    .trim();

  return fmt
    .replace(/yyyy/gi, yyyy)
    .replace(/yy/gi, yy)
    .replace(/dd/gi, dd)
    .replace(/mm/gi, mm)
    .replace(/hh/gi, hh)
    .replace(/ss/gi, '00');
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

    // --- Column widths (ExcelJS chars -> px, 1 char ≈ 7.5px) ---
    worksheet.columns.forEach((col, idx) => {
      if (col?.width > 0) {
        columnWidths[idx] = Math.round(col.width * 7.5);
      }
    });

    const totalRows = worksheet.rowCount || 1;
    let rowsDone = 0;

    // Build set of merged slave cell addresses to skip
    const slaveCells = new Set();
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

        // Mark all slave cells (everything except top-left)
        for (let ri = startCell.row; ri <= endCell.row; ri++) {
          for (let ci = startCell.col; ci <= endCell.col; ci++) {
            if (ri !== startCell.row || ci !== startCell.col) {
              slaveCells.add(`${ri}_${ci}`);
            }
          }
        }
      } catch { /* skip invalid ranges */ }
    });

    // --- Rows and cells ---
    worksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      const r = rowNum - 1;

      if (row.height > 0) {
        rowHeights[r] = Math.round(row.height * 1.333);
      }

      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const c = colNum - 1;

        // Skip slave cells of merged ranges
        if (slaveCells.has(`${rowNum}_${colNum}`)) return;

        const rawVal = cell.value;
        if (rawVal === null || rawVal === undefined) {
          // Still process if cell has styling
          if (!cell.style || Object.keys(cell.style).length === 0) return;
        }

        const cellData = {};

        // --- Value and formula ---
        if (rawVal !== null && rawVal !== undefined) {
          if (typeof rawVal === 'object' && 'formula' in rawVal) {
            // Formula cell
            const formula = rawVal.formula || rawVal.sharedFormula || '';
            const result = rawVal.result;
            cellData.f = formula.startsWith('=') ? formula.slice(1) : formula;

            if (result instanceof Date) {
              cellData.v = toExcelSerial(result);
              cellData.m = formatDate(result, cell.numFmt);
              cellData.ct = { fa: cell.numFmt || 'dd.mm.yy', t: 'd' };
            } else if (result && typeof result === 'object' && result.error) {
              cellData.v = result.error;
              cellData.m = result.error;
            } else {
              cellData.v = result ?? '';
              cellData.m = cell.text || String(result ?? '');
            }
          } else if (rawVal instanceof Date) {
            // Date cell — store as serial number with format
            cellData.v = toExcelSerial(rawVal);
            cellData.m = formatDate(rawVal, cell.numFmt);
            cellData.ct = { fa: cell.numFmt || 'dd.mm.yy', t: 'd' };
            cellData.t = 'n';
          } else if (typeof rawVal === 'object' && 'richText' in rawVal) {
            const text = rawVal.richText.map((r) => r.text).join('');
            cellData.v = text;
            cellData.m = text;
          } else {
            cellData.v = rawVal;
            cellData.m = cell.text || String(rawVal);
          }
        }

        // Cell type
        if (typeof cellData.v === 'number' && !cellData.ct) cellData.t = 'n';
        else if (typeof cellData.v === 'boolean') cellData.t = 'b';
        else if (typeof cellData.v === 'string' && cellData.v && !cellData.ct) cellData.t = 's';

        // --- Styles ---
        const fill = cell.fill;
        if (fill?.type === 'pattern' && fill.pattern !== 'none' && fill.pattern !== null) {
          const fg = fill.fgColor;
          if (fg?.argb) {
            const color = argbToCss(fg.argb);
            if (color) cellData.bg = color;
          }
        }

        const font = cell.font;
        if (font) {
          if (font.bold) cellData.bl = 1;
          if (font.italic) cellData.it = 1;
          if (font.underline) cellData.un = 1;
          if (font.size) cellData.fs = font.size;
          if (font.name) cellData.ff = font.name;
          if (font.color?.argb) {
            const fc = argbToCss(font.color.argb);
            if (fc) cellData.fc = fc;
          }
        }

        const alignment = cell.alignment;
        if (alignment) {
          if (alignment.horizontal) {
            cellData.ht = { left: 1, center: 0, right: 2 }[alignment.horizontal] ?? 1;
          }
          if (alignment.vertical) {
            cellData.vt = { middle: 0, top: 1, bottom: 2 }[alignment.vertical] ?? 0;
          }
          if (alignment.wrapText) cellData.tb = 2;
        }

        const border = cell.border;
        if (border) {
          const bd = {};
          if (border.top) bd.t = borderStyle(border.top);
          if (border.bottom) bd.b = borderStyle(border.bottom);
          if (border.left) bd.l = borderStyle(border.left);
          if (border.right) bd.r = borderStyle(border.right);
          if (Object.keys(bd).length) cellData.bd = bd;
        }

        cells[`${r}_${c}`] = cellData;
      });

      rowsDone++;
      if (onProgress) {
        const pct = Math.round(
          ((sheetsDone / totalSheets) + (rowsDone / totalRows / totalSheets)) * 95
        );
        onProgress(Math.min(pct, 94));
      }
    });

    sheets.push({ name: worksheet.name, cells, columnWidths, rowHeights, merges });
    sheetsDone++;
  }

  return sheets;
}

export async function exportExcel(sheetsData) {
  const workbook = new ExcelJS.Workbook();

  for (const sheet of sheetsData) {
    const worksheet = workbook.addWorksheet(sheet.name || 'Sheet1');

    if (sheet.columnWidths) {
      Object.entries(sheet.columnWidths).forEach(([idx, px]) => {
        worksheet.getColumn(parseInt(idx) + 1).width = Math.round((px / 7.5) * 10) / 10;
      });
    }

    if (sheet.rowHeights) {
      Object.entries(sheet.rowHeights).forEach(([rowIdx, px]) => {
        worksheet.getRow(parseInt(rowIdx) + 1).height = Math.round((px / 1.333) * 10) / 10;
      });
    }

    if (sheet.cells) {
      Object.entries(sheet.cells).forEach(([key, cd]) => {
        const [r, c] = key.split('_').map(Number);
        const cell = worksheet.getCell(r + 1, c + 1);

        if (cd.f) {
          cell.value = { formula: cd.f, result: cd.v };
        } else if (cd.ct?.t === 'd' && typeof cd.v === 'number') {
          // Convert Excel serial back to Date
          const date = new Date((cd.v - 25569) * 86400000);
          cell.value = date;
          if (cd.ct?.fa) cell.numFmt = cd.ct.fa;
        } else if (cd.v !== undefined && cd.v !== null) {
          cell.value = cd.v;
        }

        const fontStyle = {};
        if (cd.bl) fontStyle.bold = true;
        if (cd.it) fontStyle.italic = true;
        if (cd.un) fontStyle.underline = true;
        if (cd.fs) fontStyle.size = cd.fs;
        if (cd.ff) fontStyle.name = cd.ff;
        if (cd.fc) fontStyle.color = { argb: cssToArgb(cd.fc) };
        if (Object.keys(fontStyle).length) cell.font = fontStyle;

        if (cd.bg) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cssToArgb(cd.bg) } };
        }

        const alignStyle = {};
        if (cd.ht !== undefined) alignStyle.horizontal = ['center', 'left', 'right'][cd.ht] || 'left';
        if (cd.vt !== undefined) alignStyle.vertical = ['middle', 'top', 'bottom'][cd.vt] || 'middle';
        if (cd.tb === 2) alignStyle.wrapText = true;
        if (Object.keys(alignStyle).length) cell.alignment = alignStyle;

        if (cd.bd) {
          const toBorder = (b) => b ? { style: b.style === '2' ? 'medium' : 'thin', color: { argb: cssToArgb(b.color) } } : undefined;
          const border = {};
          if (cd.bd.t) border.top = toBorder(cd.bd.t);
          if (cd.bd.b) border.bottom = toBorder(cd.bd.b);
          if (cd.bd.l) border.left = toBorder(cd.bd.l);
          if (cd.bd.r) border.right = toBorder(cd.bd.r);
          cell.border = border;
        }
      });
    }

    if (sheet.merges) {
      Object.values(sheet.merges).forEach(({ r, c, rs, cs }) => {
        if (rs > 1 || cs > 1) {
          try { worksheet.mergeCells(r + 1, c + 1, r + rs, c + cs); } catch { /* skip */ }
        }
      });
    }
  }

  return workbook.xlsx.writeBuffer();
}
