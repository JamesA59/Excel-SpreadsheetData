# Manipulating Excel Files With Openpxyl Challenge:

# Split a single worksheet into multiple worksheets

import openpyxl
from openpyxl.utils.cell import column_index_from_string

filename = "FinancialSample.xlsx"
wb = openpyxl.load_workbook(filename)

sheet = wb.active
column = sheet["B"]
all = [column[x].value for x in range(len(column))]
countries = list(set(all))
countries.sort()
countries.__delitem__(1)

row = sheet[1]
header = [row[x].value for x in range(len(row))]

t = 2
all_data = []
while t <= 701:

    rows = sheet[t]
    data = [rows[x].value for x in range(len(rows))]

    all_data.append(data)
    t += 1

y = 3
n = str(y)
p = 0
new = "sheet" + n

filters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]

for i in range(len(countries)):
    new = wb.create_sheet(countries[p])
    new.append(header)

    for sheet_name in new:
        sheet = wb[countries[p]]
        filters = sheet.auto_filter
        filters.ref = sheet.dimensions

    for letter in filters:
        letter = str(letter)
        tile = letter + "1"
        tile = float(tile)
        new[tile].style = "Accent 2"
    
    for row in all_data:
        if countries[p] in row:
            new.append(row)

    y += 1
    p += 1

wb.save("newFinancialSample.xlsx")
print("Workbook created successfully!")