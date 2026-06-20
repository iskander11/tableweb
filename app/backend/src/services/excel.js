import ExcelJS from 'exceljs';

// Convert ARGB hex to CSS color
const argbToCss = (argb) => {
  if (!argb || argb === '00000000') return null;
  if (argb.length === 8) return `#${argb.slice(2)}`;
  return `#${argb}`;
};

// Convert CSS hex to ARGB
const cssToArgb = (css) => {
  if (!css) return '00000000';
  const hex = css.replace('#', '');
  return `FF${hex.toUpperCase()}`;
};

export async function importExcel(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheets = [];

  workbook.eachSheet((worksheet) => {
    const cells = {};
    const columnWidths = {};
    const rowHeights = {};

    // Column widths
    worksheet.columns.forEach((col, idx) => {
      if (col.width) columnWidths[idx] = Math.round(col.width * 8); // pts to px approx
    });

    // Row heights
    worksheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
      if (row.height) rowHeights[rowNum - 1] = row.height * 1.33; // pts to px

      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        const r = rowNum - 1;
        const c = colNum - 1;
        const key = `${r}_${c}`;

        const cellData = {
          v: cell.value,
          m: cell.text || '',
        };

        // Formula
        if (cell.formula) {
          cellData.f = cell.formula;
          cellData.v = cell.result ?? '';
        }

        // Style
        const style = {};
        if (cell.font) {
          if (cell.font.bold) style.bl = 1;
          if (cell.font.italic) style.it = 1;
          if (cell.font.size) style.fs = cell.font.size;
          if (cell.font.color?.argb) style.fc = argbToCss(cell.font.color.argb);
        }
        if (cell.fill?.fgColor?.argb) {
          style.bg = argbToCss(cell.fill.fgColor.argb);
        }
        if (cell.alignment) {
          if (cell.alignment.horizontal) style.ht = { left: 1, center: 0, right: 2 }[cell.alignment.horizontal] ?? 1;
          if (cell.alignment.wrapText) style.tb = 2;
        }
        if (cell.border) {
          style.bd = {
            t: cell.border.top ? { style: 'thin', color: '#000000' } : null,
            b: cell.border.bottom ? { style: 'thin', color: '#000000' } : null,
            l: cell.border.left ? { style: 'thin', color: '#000000' } : null,
            r: cell.border.right ? { style: 'thin', color: '#000000' } : null,
          };
        }

        if (Object.keys(style).length) cellData.s = style;

        // Merge info
        worksheet.model.merges?.forEach((merge) => {
          // handled by FortuneSheet merge map
        });

        cells[key] = cellData;
      });
    });

    // Merged cells
    const merges = {};
    (worksheet.model.merges || []).forEach((mergeStr) => {
      const [start, end] = mergeStr.split(':');
      const sc = worksheet.getCell(start);
      const ec = worksheet.getCell(end);
      const r = sc.row - 1;
      const c = sc.col - 1;
      merges[`${r}_${c}`] = {
        r, c,
        rs: ec.row - sc.row + 1,
        cs: ec.col - sc.col + 1,
      };
    });

    sheets.push({
      name: worksheet.name,
      cells,
      columnWidths,
      rowHeights,
      merges,
    });
  });

  return sheets;
}

export async function exportExcel(sheetsData) {
  const workbook = new ExcelJS.Workbook();

  for (const sheet of sheetsData) {
    const worksheet = workbook.addWorksheet(sheet.name || 'Sheet1');

    // Apply column widths
    if (sheet.columnWidths) {
      Object.entries(sheet.columnWidths).forEach(([idx, px]) => {
        worksheet.getColumn(parseInt(idx) + 1).width = px / 8;
      });
    }

    // Apply row heights
    if (sheet.rowHeights) {
      Object.entries(sheet.rowHeights).forEach(([rowIdx, px]) => {
        worksheet.getRow(parseInt(rowIdx) + 1).height = px / 1.33;
      });
    }

    // Apply cell data
    if (sheet.cells) {
      Object.entries(sheet.cells).forEach(([key, cellData]) => {
        const [r, c] = key.split('_').map(Number);
        const cell = worksheet.getCell(r + 1, c + 1);

        if (cellData.f) {
          cell.value = { formula: cellData.f, result: cellData.v };
        } else if (cellData.v !== undefined && cellData.v !== null) {
          cell.value = cellData.v;
        }

        if (cellData.s) {
          const s = cellData.s;
          cell.font = {
            bold: !!s.bl,
            italic: !!s.it,
            size: s.fs || 11,
            color: s.fc ? { argb: cssToArgb(s.fc) } : undefined,
          };
          if (s.bg) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: cssToArgb(s.bg) },
            };
          }
          if (s.ht !== undefined) {
            cell.alignment = {
              horizontal: ['center', 'left', 'right'][s.ht] || 'left',
              wrapText: s.tb === 2,
            };
          }
        }
      });
    }

    // Apply merges
    if (sheet.merges) {
      Object.values(sheet.merges).forEach(({ r, c, rs, cs }) => {
        if (rs > 1 || cs > 1) {
          worksheet.mergeCells(r + 1, c + 1, r + rs, c + cs);
        }
      });
    }
  }

  return workbook.xlsx.writeBuffer();
}
