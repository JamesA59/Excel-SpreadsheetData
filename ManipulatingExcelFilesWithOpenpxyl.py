# Manipulating Excel Files with openpxyl:

# Overview of openpxyl:

# The openpyxl library:
#       Free, open source Python library for working with Excel files (.xlsx)
#       Read, write, and manipulate the content of workbooks and worksheets
#       Userful for automating many common tasks:
#           Create spreadsheets from other data sources
#           Combine multiple data sources and generate formatted reports
#           Create charts and templates programmatically
#       Build and deploy powerful data processing pipelines and workflows
# Openpyxl is also used by the Pandas library to perform operations on Excel files


# Loading and exploring a workbook:

# Open, load, and explore workbook content 

'''
import openpyxl

filename = "FinancialSample.xlsx"

# Load the workbook
workbook = openpyxl.load_workbook(filename)

# Print basic information
print(f"Number of worksheets: {len(workbook.sheetnames)}")
for worksheet_name in workbook.sheetnames:
    worksheet = workbook[worksheet_name]
    print(f"\nWorksheet: {worksheet_name}")

# Explore each worksheet

    
# Get dimensions
    dimensions = worksheet.dimensions
    print(f"  - Dimensions: {dimensions}")

    print(f"Min row: {worksheet.min_row}")
    print(f"Max row: {worksheet.max_row}")
    print(f"Min column: {worksheet.min_column}")
    print(f"Max column: {worksheet.max_column}")

# Check if the worksheet is empty
    if worksheet.max_row == 1 and worksheet.max_column == 1:
        print(f"  - Worksheet is empty")
    else:
        cell = worksheet["A1"]
        print(f"  - Top=left cell calue: {cell.value}")
        cell = worksheet.cell(row=worksheet.max_row, column=worksheet.max_column)
        print(f"  - Bottom-right cell value: {cell.value}")
'''


# Creating a workbook:

# Create a new workbook with worksheets and add content 

from openpyxl import Workbook
import datetime
import random

# Create a new workbook
wb = Workbook()

# Get the active worksheet and name it "First"
sheet = wb.active
sheet.title = "First"

# Add some data to the new sheet
sheet["A1"] = "Test Data"
sheet["B1"] = 123.4567
sheet["C1"] = datetime.datetime(2030, 4, 1)

# Use the cell() function to fill a row with values
for i in range(1, 11):
    sheet.cell(row=5, column=i).value = random.randint(1,50)

# Create a second worksheet
sheet2 = wb.create_sheet("Second")
sheet2.cell(row=2, column=2).value = "More Data"

# Use the append() function to add rows to the end of the sheet
sheet2.append(["One", "Two", "Three"])
sheet2.append(["One", "Two", "Three"])
sheet2.append(["One", "Two", "Three"])

# Save the workbook - values don't update until we do this!
wb.save("NewWorkbook.xlsx")
print("Workbook has been created successfully!")