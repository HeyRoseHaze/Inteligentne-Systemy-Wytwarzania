import openpyxl as op
from openpyxl.styles import Alignment

wb = op.Workbook()
ws = wb.active
ws.title = 'Harmonogram'
ws['A1'] = 42

for i in range(101):
    column = i + 2
    cell = ws.cell(row = 10, column = column, value = i)
    cell.alignment = Alignment(horizontal = 'center')
    ws.column_dimensions[op.utils.cell.get_column_letter(column)].width = 3

wb.save("harmonogram.xlsx")
