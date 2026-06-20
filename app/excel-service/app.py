from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
import io
import json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB


def argb_to_css(argb):
    """Convert ARGB string (FFRRGGBB) to CSS hex (#RRGGBB)"""
    if not argb or len(argb) < 6:
        return None
    # openpyxl returns 8-char ARGB
    if len(argb) == 8:
        alpha = argb[:2]
        rgb = argb[2:]
    else:
        rgb = argb
    # Skip fully transparent or black (00000000)
    if rgb == '000000' or argb == '00000000':
        return None
    return f'#{rgb.upper()}'


def css_to_argb(css):
    """Convert CSS hex (#RRGGBB) to ARGB (FFRRGGBB)"""
    if not css:
        return 'FF000000'
    return f'FF{css.lstrip("#").upper()}'


def get_border_style(side):
    if not side or not side.border_style:
        return None
    style = '2' if side.border_style in ('medium', 'thick') else '1'
    color = '#000000'
    if side.color and side.color.rgb and side.color.rgb != '00000000':
        color = f'#{side.color.rgb[2:]}'
    return {'style': style, 'color': color}


def cell_value_and_type(cell):
    """Extract value, display text, formula, and type from a cell"""
    val = cell.value
    formula = None
    display = ''

    if val is None:
        return None, None, None, ''

    # Formula
    if isinstance(val, str) and val.startswith('='):
        formula = val[1:]  # strip leading '='
        val = None         # we don't have the computed result
        display = ''
        return val, formula, None, display

    # Date/datetime
    if cell.is_date and val is not None:
        from datetime import datetime, date
        if isinstance(val, (datetime, date)):
            # Convert to Excel serial number
            from openpyxl.utils.datetime import to_excel
            serial = to_excel(val)
            # Format for display
            fmt = (cell.number_format or '').split(';')[0]
            if isinstance(val, datetime):
                display = val.strftime('%d.%m.%y') if 'yy' in fmt.lower() else str(val.date())
            else:
                display = val.strftime('%d.%m.%y')
            return serial, None, 'd', display

    # Number
    if isinstance(val, (int, float)):
        display = str(val)
        return val, None, 'n', display

    # Boolean
    if isinstance(val, bool):
        return val, None, 'b', str(val)

    # String
    display = str(val)
    return val, None, 's', display


@app.route('/parse', methods=['POST'])
def parse_excel():
    """Parse an Excel file and return sheet data in FortuneSheet format"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    content = file.read()

    try:
        wb = load_workbook(
            io.BytesIO(content),
            data_only=True,   # get computed values instead of formulas
            rich_text=False
        )
        # Also load with formulas to get formula strings
        wb_formulas = load_workbook(
            io.BytesIO(content),
            data_only=False,
            rich_text=False
        )
    except Exception as e:
        return jsonify({'error': f'Failed to load workbook: {str(e)}'}), 400

    sheets = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws_f = wb_formulas[sheet_name]

        cells = {}
        column_widths = {}
        row_heights = {}
        merges = {}

        # --- Column widths (Excel chars → px, 1 char ≈ 7.5px) ---
        for col_letter, col_dim in ws.column_dimensions.items():
            if col_dim.width and col_dim.width > 0:
                col_idx = column_index_from_string(col_letter) - 1  # 0-based
                column_widths[col_idx] = round(col_dim.width * 7.5)

        # --- Row heights (pt → px, 1pt ≈ 1.333px) ---
        for row_idx, row_dim in ws.row_dimensions.items():
            if row_dim.height and row_dim.height > 0:
                row_heights[row_idx - 1] = round(row_dim.height * 1.333)

        # --- Merged cells ---
        slave_cells = set()
        for merge_range in ws.merged_cells.ranges:
            r = merge_range.min_row - 1   # 0-based
            c = merge_range.min_col - 1
            rs = merge_range.max_row - merge_range.min_row + 1
            cs = merge_range.max_col - merge_range.min_col + 1

            if rs > 1 or cs > 1:
                merges[f'{r}_{c}'] = {'r': r, 'c': c, 'rs': rs, 'cs': cs}

            # Mark slave cells
            for ri in range(merge_range.min_row, merge_range.max_row + 1):
                for ci in range(merge_range.min_col, merge_range.max_col + 1):
                    if ri != merge_range.min_row or ci != merge_range.min_col:
                        slave_cells.add((ri, ci))

        # --- Cells ---
        for row in ws.iter_rows():
            for cell in row:
                ri = cell.row
                ci = cell.column

                # Skip slave cells of merged ranges
                if (ri, ci) in slave_cells:
                    continue

                r = ri - 1  # 0-based
                c = ci - 1

                # Get formula from formula workbook
                formula_cell = ws_f.cell(row=ri, column=ci)
                formula_val = formula_cell.value
                has_formula = isinstance(formula_val, str) and formula_val.startswith('=')

                val, formula, cell_type, display = cell_value_and_type(cell)

                if has_formula:
                    formula = formula_val[1:]  # strip '='
                    # val stays as the computed result from data_only workbook

                # Skip completely empty cells with no style
                has_style = False
                if cell.fill and isinstance(cell.fill, PatternFill) and cell.fill.fill_type not in (None, 'none'):
                    has_style = True
                if cell.font and (cell.font.bold or cell.font.italic or cell.font.color):
                    has_style = True
                if cell.border:
                    b = cell.border
                    if b.top or b.bottom or b.left or b.right:
                        has_style = True

                if val is None and not formula and not has_style:
                    continue

                cell_data = {}

                # Value
                if val is not None:
                    cell_data['v'] = val
                if display:
                    cell_data['m'] = display
                if formula:
                    cell_data['f'] = formula
                if cell_type:
                    if cell_type == 'd':
                        cell_data['ct'] = {'fa': (cell.number_format or '').split(';')[0], 't': 'd'}
                        cell_data['t'] = 'n'
                    else:
                        cell_data['t'] = cell_type

                # --- Background color ---
                fill = cell.fill
                if fill and isinstance(fill, PatternFill) and fill.fill_type not in (None, 'none'):
                    fg = fill.fgColor
                    if fg:
                        color = None
                        if fg.type == 'rgb':
                            color = argb_to_css(fg.rgb)
                        elif fg.type == 'indexed':
                            # Common indexed colors
                            indexed_colors = {
                                0: None, 1: '#000000', 2: '#FFFFFF', 3: '#FF0000',
                                4: '#00FF00', 5: '#0000FF', 6: '#FFFF00', 7: '#FF00FF',
                                8: '#00FFFF', 9: '#000000', 10: '#FFFFFF',
                                64: None,  # transparent
                            }
                            color = indexed_colors.get(fg.indexed)
                        if color:
                            cell_data['bg'] = color

                # --- Font ---
                font = cell.font
                if font:
                    if font.bold:
                        cell_data['bl'] = 1
                    if font.italic:
                        cell_data['it'] = 1
                    if font.underline:
                        cell_data['un'] = 1
                    if font.size:
                        cell_data['fs'] = font.size
                    if font.name:
                        cell_data['ff'] = font.name
                    if font.color:
                        fc = None
                        if font.color.type == 'rgb':
                            fc = argb_to_css(font.color.rgb)
                        if fc:
                            cell_data['fc'] = fc

                # --- Alignment ---
                alignment = cell.alignment
                if alignment:
                    if alignment.horizontal:
                        ht_map = {'left': 1, 'center': 0, 'right': 2, 'general': 1}
                        cell_data['ht'] = ht_map.get(alignment.horizontal, 1)
                    if alignment.vertical:
                        vt_map = {'center': 0, 'top': 1, 'bottom': 2}
                        cell_data['vt'] = vt_map.get(alignment.vertical, 0)
                    if alignment.wrap_text:
                        cell_data['tb'] = 2

                # --- Borders ---
                border = cell.border
                if border:
                    bd = {}
                    t = get_border_style(border.top)
                    b = get_border_style(border.bottom)
                    l = get_border_style(border.left)
                    rr = get_border_style(border.right)
                    if t: bd['t'] = t
                    if b: bd['b'] = b
                    if l: bd['l'] = l
                    if rr: bd['r'] = rr
                    if bd:
                        cell_data['bd'] = bd

                cells[f'{r}_{c}'] = cell_data

        sheets.append({
            'name': sheet_name,
            'cells': cells,
            'columnWidths': column_widths,
            'rowHeights': row_heights,
            'merges': merges,
        })

    return jsonify({'sheets': sheets})


@app.route('/export', methods=['POST'])
def export_excel():
    """Convert FortuneSheet data back to Excel file"""
    data = request.get_json()
    sheets_data = data.get('sheets', [])

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    for sheet in sheets_data:
        ws = wb.create_sheet(title=sheet.get('name', 'Sheet1'))

        # Column widths (px → chars)
        for col_idx_str, px in (sheet.get('columnWidths') or {}).items():
            col_letter = get_column_letter(int(col_idx_str) + 1)
            ws.column_dimensions[col_letter].width = round(px / 7.5, 1)

        # Row heights (px → pt)
        for row_idx_str, px in (sheet.get('rowHeights') or {}).items():
            ws.row_dimensions[int(row_idx_str) + 1].height = round(px / 1.333, 1)

        # Cells
        for key, cd in (sheet.get('cells') or {}).items():
            r_str, c_str = key.split('_')
            r = int(r_str) + 1
            c = int(c_str) + 1
            cell = ws.cell(row=r, column=c)

            # Value / formula
            if cd.get('f'):
                cell.value = f"={cd['f']}"
            elif cd.get('ct', {}).get('t') == 'd' and cd.get('v') is not None:
                from openpyxl.utils.datetime import from_excel
                try:
                    cell.value = from_excel(cd['v'])
                    fmt = cd.get('ct', {}).get('fa', 'DD.MM.YY')
                    cell.number_format = fmt
                except Exception:
                    cell.value = cd.get('m', '')
            elif cd.get('v') is not None:
                cell.value = cd['v']

            # Background
            if cd.get('bg'):
                argb = css_to_argb(cd['bg'])
                cell.fill = PatternFill(fill_type='solid', fgColor=argb)

            # Font
            font_kwargs = {}
            if cd.get('bl'): font_kwargs['bold'] = True
            if cd.get('it'): font_kwargs['italic'] = True
            if cd.get('un'): font_kwargs['underline'] = 'single'
            if cd.get('fs'): font_kwargs['size'] = cd['fs']
            if cd.get('ff'): font_kwargs['name'] = cd['ff']
            if cd.get('fc'): font_kwargs['color'] = css_to_argb(cd['fc'])
            if font_kwargs:
                cell.font = Font(**font_kwargs)

            # Alignment
            align_kwargs = {}
            if cd.get('ht') is not None:
                align_kwargs['horizontal'] = ['center', 'left', 'right'][cd['ht']]
            if cd.get('vt') is not None:
                align_kwargs['vertical'] = ['center', 'top', 'bottom'][cd['vt']]
            if cd.get('tb') == 2:
                align_kwargs['wrap_text'] = True
            if align_kwargs:
                cell.alignment = Alignment(**align_kwargs)

            # Borders
            if cd.get('bd'):
                def make_side(b):
                    if not b:
                        return Side()
                    return Side(
                        border_style='medium' if b.get('style') == '2' else 'thin',
                        color=css_to_argb(b.get('color', '#000000'))
                    )
                bd = cd['bd']
                cell.border = Border(
                    top=make_side(bd.get('t')),
                    bottom=make_side(bd.get('b')),
                    left=make_side(bd.get('l')),
                    right=make_side(bd.get('r')),
                )

        # Merged cells
        for key, m in (sheet.get('merges') or {}).items():
            if m['rs'] > 1 or m['cs'] > 1:
                try:
                    ws.merge_cells(
                        start_row=m['r'] + 1, start_column=m['c'] + 1,
                        end_row=m['r'] + m['rs'], end_column=m['c'] + m['cs']
                    )
                except Exception:
                    pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='export.xlsx'
    )


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=False)
