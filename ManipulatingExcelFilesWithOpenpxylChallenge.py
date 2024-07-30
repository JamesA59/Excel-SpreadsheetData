# Manipulating Excel Files With Openpxyl Challenge:

# Split a single worksheet into multiple worksheets

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import pandas as pd

filename = "FinancialSample.xlsx"
workbook = openpyxl.load_workbook(filename)

sheet = workbook.active
column = sheet["B"]
all = [column[x].value for x in range(len(column))]
countries = list(set(all))
countries.sort()
countries.__delitem__(1)

row = sheet[1]
header = [row[x].value for x in range (len(row))]
print(header)

y = 3
n = str(y)
p = 0
new = "sheet" + n

for i in range(len(countries)):
    new = workbook.create_sheet(countries[p])
    new.append(header)
     #sheet["A1:P1"].style = "Accent 2"
    t = 2
    rows = sheet[t]
    '''
    if rows[1].value == countries[p]:
        data = [rows[x].value for x in range (len(rows))]
        new.append(data)
        t += 1
    else:
        t += 1
    '''
    r=str(t)
    cell = "B" + r

    if new[cell] == countries[p]:
        data = [rows[x].value for x in range (len(rows))]
        new.append(data)
        t += 1
    else:
        t += 1

    y += 1
    p += 1

workbook.save("newFinancialSample.xlsx")
print("Workbook created successfully!")