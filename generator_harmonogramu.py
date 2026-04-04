import openpyxl as op


wb = op.Workbook()
ws = wb.active
ws.title = 'Harmonogram'
ws['A1'] = 42

for i in range(101):
    column = i + 2
    ws.cell(row = 10, column = column, value = i)

wb.save("harmonogram.xlsx")
