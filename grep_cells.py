import os
import sys
import textwrap
import openpyxl

spreadsheet_path = 'copied'
if len(sys.argv) > 1:
    grep_values = sys.argv[1:]
else:
    grep_values = ('FWS', )


print('Grepping {!r} for: {}'
      ''.format(os.path.abspath(spreadsheet_path), grep_values))

for fn in os.listdir(spreadsheet_path):
    try:
        wb = openpyxl.load_workbook(os.path.join(spreadsheet_path, fn))
    except Exception as ex:
        continue

    for ws in wb.worksheets:
        found = any(to_grep in str(cell.value)
                    for row in ws.rows
                    for cell in row
                    if cell.value is not None
                    for to_grep in grep_values
                    )

        if found:
            header = '{}: {}'.format(fn, ws.title)
            print()
            print(header)
            print('-' * len(header))

            for row in ws.rows:
                for cell in row:
                    if cell.value is not None:
                        for to_grep in grep_values:
                            if to_grep in str(cell.value):
                                print(cell)
                                print(textwrap.indent(cell.value, '    '))
            print()
