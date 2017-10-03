import os
import openpyxl


def copy_worksheet(ws, new_ws):
    '''Destructively copy a worksheet to a new one'''
    for row in ws.rows:
        for cell in row:
            new_cell = new_ws[cell.coordinate]
            new_cell.value = cell.value
            new_cell.style = cell.style
            new_cell.border = cell.border.copy()
            new_cell.font = cell.font.copy()
            new_cell.fill = cell.fill.copy()
            new_cell.alignment = cell.alignment.copy()
            new_cell.number_format = cell.number_format

    for merged_range in ws.merged_cell_ranges:
        new_ws.merge_cells(merged_range)
        first = merged_range.split(':')[0]
        orig, merged = ws[first], new_ws[first]
        merged.value = orig.value
        cell = ws[first]
        for row in new_ws[merged_range]:
            for new_cell in row:
                new_cell.style = cell.style
                new_cell.border = cell.border.copy()
                new_cell.font = cell.font.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.alignment = cell.alignment.copy()
                new_cell.number_format = cell.number_format

    for row_id, row_dim in ws.row_dimensions.items():
        new_row = new_ws.row_dimensions[row_id]
        new_row.height = row_dim.height

    for col_id, col_dim in ws.column_dimensions.items():
        new_column = new_ws.column_dimensions[col_id]
        new_column.width = col_dim.width


new_wb = openpyxl.Workbook()

spreadsheet_path = 'copied'
grep_values = ('FWS', 'Lauer')
matched_sheets = 0

for fn in os.listdir(spreadsheet_path):
    try:
        wb = openpyxl.load_workbook(os.path.join(spreadsheet_path, fn))
    except Exception as ex:
        print('Failed to load: {}'.format(fn), ex)
        continue

    for ws in wb.worksheets:
        print('File: {} worksheet: {}'.format(fn, ws))
        found = any(to_grep in str(cell.value)
                    for row in ws.rows
                    for cell in row
                    if cell.value is not None
                    for to_grep in grep_values
                    )

        if not found:
            continue

        print('Match')
        matched_sheets += 1

        # title = '{}_{}'.format(fn.rsplit('.', 1)[0], ws.title)
        title = ws.title
        new_ws = new_wb.create_sheet(title)
        copy_worksheet(ws, new_ws)

        new_ws.print_area = new_ws.dimensions
        new_ws.print_options.horizontalCentered = True
        new_ws.print_options.verticalCentered = True
        header_text = 'File: {} Sheet: {}'.format(fn.rsplit('.', 1)[0],
                                                  ws.title)
        new_ws.firstHeader.center.text = header_text
        new_ws.page_setup.fitToPage = True

    # break

new_wb.save('test.xlsx')
print('Matching sheet count: ', matched_sheets)
