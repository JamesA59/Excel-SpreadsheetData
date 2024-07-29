# Manipulating Excel Files With Openpxyl Challenge:

# Split a single worksheet into multiple worksheets

import openpyxl

filename = "FinancialSample.xlsx"
wb = openpyxl.load_workbook(filename)

sheet = wb.active
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
    new = wb.create_sheet(countries[p])
    new.append(header)
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