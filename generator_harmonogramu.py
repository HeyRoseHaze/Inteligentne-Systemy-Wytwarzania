import openpyxl as op
from openpyxl.styles import Alignment, PatternFill
import random

wb = op.Workbook()
ws = wb.active
ws.title = 'Harmonogram'

for i in range(101):
    column = i + 2
    cell = ws.cell(row = 10, column = column, value = i)
    cell.alignment = Alignment(horizontal = 'center')
    ws.column_dimensions[op.utils.cell.get_column_letter(column)].width = 3


colors = [
    "FF0000", # Czerwony
    "FFFF00", # Żółty
    "00DBFF", # jasnoniebieski
    "00FF11", # jasnozielony
    "004EFF", # ciemnoniebieski
    "BA4EFF", # fioletowy
    "858B83", # szary
    "008100", # ciemnozielony
    "F88100", # pomarańczowy
]

fills = [PatternFill(start_color = color, end_color= color, fill_type='solid') for color in colors]

# Macierz C
C = [
    [10, 20, 30, 40, 50, 60, 70, 80, 100],
    [11, 21, 31, 41, 51, 61, 71, 81, 100],
    [12, 22, 32, 42, 52, 62, 72, 82, 100],
    [13, 23, 33, 43, 53, 63, 73, 83, 100],
    [14, 24, 34, 44, 54, 64, 74, 84, 100],
    [15, 25, 35, 45, 55, 65, 75, 85, 100],
    [16, 26, 36, 46, 56, 66, 76, 86, 100],
    [17, 27, 37, 47, 57, 67, 77, 87, 100]
]

zadania = list(range(9))
random.shuffle(zadania) 

for i in task:
    nr_maszyny = i[0]
    nr_zadania = i[1]
    start = i[2]
    stop = i[3]

    wiersz = nr_maszyny + 1
    kolor = nr_zadania - 1
    malowanie = fills[kolor]

    długosc = stop - start

    for t in range(długosc):
        kolumna = start + t + 2
        ws.cell(row=wiersz, column=kolumna).fill = malowanie

wb.save("harmonogram.xlsx")