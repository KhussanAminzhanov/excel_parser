# w, h = 10, 10
# matrix = [[0 for x in range(w)] for y in range(h)]
# for r in range(h):
#     for i in range(w):
#         matrix[r][i] = r + i

# for r in range(h):
#     for i in range(w):
#         print(matrix[r][i], end = "\t")
#     print()
from os import sep
import openpyxl

w, h = 16, 20
wb = openpyxl.load_workbook("workbook.xlsx")
sheet = wb['Sheet2']
table = [[0 for x in range(w)] for y in range(h)]

def getRow(start, end, rowNumber):
    row = [0 for i in range(end - start + 1)]
    for i in range(start, end + 1):
        c = sheet[chr(i) + str(rowNumber)].value
        if (c == None):
            c = 0
        row[i - start] = int(c)
    return row

for rowNumber in range(8, 28):
    row = getRow(68,83,rowNumber)
    for i in range(0, len(row)):
        table[rowNumber - 8][i] = row[i]
    

for i in range(20):
    print(table[i])
