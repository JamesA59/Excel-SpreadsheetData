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

for worksheet_name in workbook.sheetnames:
    worksheet = workbook[worksheet_name]
    dimensions = workbook.dimensions
    print(f"Min row: {worksheet.min_row}")



y = 3
n = str(y)
p = 0
new = "sheet" + n
#t = 2
#rows = sheet[t]


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
    #for i in range(len())


    y += 1
    p += 1



'''
sheet3 = wb.create_sheet(countries[0])
sheet3.append(header)

sheet4 = wb.create_sheet(countries[1])
sheet4.append(header)

sheet5 = wb.create_sheet(countries[2])
sheet5.append(header)

sheet6 = wb.create_sheet(countries[3])
sheet6.append(header)

sheet7 = wb.create_sheet(countries[4])
sheet7.append(header)
'''

wb.save("newFinancialSample.xlsx")
print("Workbook created successfully!")