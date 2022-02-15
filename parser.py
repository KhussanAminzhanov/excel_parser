import openpyxl

wb = openpyxl.load_workbook("workbook.xlsx")
copy = openpyxl.load_workbook("copy.xlsx")
w, h = 16, 20
start, end = ord('U'), ord('W')
table = [[0 for x in range(w)] for y in range(h)]

def getRow(start, end, rowNumber):
    row = [0 for i in range(end - start + 1)]
    for i in range(start, end):
        c = sheet[chr(i) + str(rowNumber)].value
        if (c == None):
            c = 0
        row[i - start] = int(c)
    return row


for s in range(2,12):
    sheet = wb['Sheet' + str(s)]
    for rowNumber in range(8,28):
        row = getRow(start, end + 1, rowNumber)
        for i in range(0, len(row)):
            table[rowNumber - 8][i] += row[i] 

for i in range(20):
    for j in (table[i]):
        print(j, end='\t')
    print()

cSheet = copy['Sheet12']
for rowNumber in range(8, 28):
    for columnNumber in range(start, end + 1):
        cSheet[chr(columnNumber) + str(rowNumber)].value = table[rowNumber - 8][columnNumber - start]

copy.save("copy_1.xlsx")
