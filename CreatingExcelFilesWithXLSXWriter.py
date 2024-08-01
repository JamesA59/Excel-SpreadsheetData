# Creating Excel Files With XLSX Writer:


# Introduction to Xlsx Writer:

# Unlike Openpyxl, this package is only used to write Excel spreadsheets-
#       You  can't read or manipulate spreadsheets
# Will learn how to use xlsx writer to create workbooks, style their contents, apply conditional formatting and filters, and create Excel tables


# Creating a workbook:


# Will create a workbook and add content to it
'''
import xlsxwriter
import datetime

# create the workbook and add a worksheet
workbook = xlsxwriter.Workbook("XlsxBasics.xlsx")
worksheet = workbook.add_worksheet("Test Sheet")

# Use Letter/Row notation
worksheet.write("A1", "Hello World")

# Use Row,Col notation
worksheet.write(1, 0, "Hello World")

# There are specific write() functions for different data types
# Python tries to guess the data type for Excel, ex. str, integer, float, etc. 
# Buy sometimes gets it wrong, so you can specify the data type
worksheet.write_number(2, 0, 12345)
worksheet.write_boolean(3, 0, True)
worksheet.write_url(4, 0, "https://www.python.org")

# Write a datetime
date_time = datetime.datetime.strptime('2030-07-28', '%Y-%m-%d')
date_format = workbook.add_format({'num_format' : 'd mmm yyyy'})
worksheet.write_datetime(5, 0, date_time, date_format)

# write multiple values into rows and columns
values = ["Good", "Morning", "Excel"]
worksheet.write_row("A6", values)
worksheet.write_column("D1", values)

# set the zoom on the sheet
worksheet.set_zoom(200)

# save the workbook
workbook.close()
'''


# Formatting Worksheet Content:

# XlsxWriter formatting
# Data is oringally in plain text format without any styling
'''
import xlsxwriter

# Sample data
data = [
    ["Item Name", "Category", "Quantity", "Wholesale Price", "Consumer Price"],
    ["Apple", "Fruits", 100, 0.50, 0.75],
    ["Banana", "Fruits", 150, 0.35, 0.50],
    ["Orange", "Fruits", 120, 0.45, 0.65],
    ["Grapes", "Fruits", 80, 0.60, 0.85],
    ["Strawberries", "Fruits", 90, 1.20, 1.50]
]

# create the workbook
workbook = xlsxwriter.Workbook('Inventory.xlsx')
worksheet = workbook.add_worksheet("Inventory")

# Use the add_format function to define formats that you can use later
# in the worksheet. NOTE: If you change the format then ALL prior instances
# of the format will be saved as the most recent one
fmt_bold = workbook.add_format({'bold':True})
fmt_money = workbook.add_format({
    "font_color" : "green",
    "num_format" : "$#,##0.00"
})

# write the data into the workbook
worksheet.write_row(0, 0, data[0], fmt_bold)
for row, itemlist in enumerate(data[1:], start=1):
    worksheet.write(row, 0, itemlist[0])
    worksheet.write(row, 1, itemlist[1], fmt_bold)
    worksheet.write(row, 2, itemlist[2])
    worksheet.write(row, 3, itemlist[3], fmt_money)
    worksheet.write(row, 4, itemlist[4], fmt_money)

worksheet.autofit()
worksheet.set_zoom(200)

workbook.close()
'''


# Creating an Excel Table:

# XlsxWriter Excel Tables
'''
import xlsxwriter

# Sample data
data = [
    ["Item Name", "Category", "Quantity", "Wholesale Price", "Consumer Price"],
    ["Apple", "Fruits", 100, 0.50, 0.75],
    ["Banana", "Fruits", 150, 0.35, 0.50],
    ["Orange", "Fruits", 120, 0.45, 0.65],
    ["Grapes", "Fruits", 80, 0.60, 0.85],
    ["Strawberries", "Fruits", 90, 1.20, 1.50]
]

# create the workbook
workbook = xlsxwriter.Workbook('Tables.xlsx')
worksheet = workbook.add_worksheet("Inventory")

fmt_bold = workbook.add_format({'bold': True})
fmt_money = workbook.add_format(
    {'font_color': 'green', 'num_format': '$#,##0.00'})

# write the data into the workbook
worksheet.write_row(0, 0, data[0], fmt_bold)
for row, itemlist in enumerate(data[1:], start=1):
    # worksheet.write_row(row, 0, itemlist)
    worksheet.write(row, 0, itemlist[0])
    worksheet.write(row, 1, itemlist[1], fmt_bold)
    worksheet.write(row, 2, itemlist[2])
    worksheet.write(row, 3, itemlist[3], fmt_money)
    worksheet.write(row, 4, itemlist[4], fmt_money)

# define a table for the worksheet
table_options = {
    "name": "InventoryData",
    "autofilter": True,
    "style": "Table Style Light 19",
    # After adding style, decided not to use banded_rows anymore
    #"banded_rows": False,
    "first_column": True,
    "columns": [
        {"header": data[0][0]},
        {"header": data[0][1]},
        {"header": data[0][2]},
        {"header": data[0][3]},
        {"header": data[0][4]},
    ]
}
worksheet.add_table("A1:E6", table_options)

worksheet.set_zoom(200)
worksheet.autofit()

workbook.close()
'''


# Apply Conditional Formatting:

# Conditional formatting is a way to specify a set of conditions under which to apply a format to a set of cells

# XlsxWriter formulas and conditional formatting
'''
import csv
import xlsxwriter


def read_csv_to_array(filename):
    # define the array that will hold the data
    data = []
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data.append(row)
    return data


# Read the data into an array of arrays
inventory_data = read_csv_to_array("Inventory.csv")

# create the workbook
workbook = xlsxwriter.Workbook('Conditional.xlsx')
worksheet = workbook.add_worksheet("Inventory")

fmt_bold = workbook.add_format({'bold': True})
fmt_money = workbook.add_format(
    {'font_color': 'green', 'num_format': '$#,##0.00'})
# define the format for the conditional expression
fmt_cond = workbook.add_format({"bg_color": "#AAFFAA", "bold": True})

# write the data into the workbook
worksheet.write_row(0, 0, inventory_data[0], fmt_bold)
# add the new header for the margin
worksheet.write(0, 5, "Margin", fmt_bold)

# add the data to the worksheet
for row, itemlist in enumerate(inventory_data[1:], start=1):
    # worksheet.write_row(row, 0, itemlist)
    worksheet.write(row, 0, itemlist[0])
    worksheet.write(row, 1, itemlist[1], fmt_bold)
    worksheet.write_number(row, 2, int(itemlist[2]))
    worksheet.write_number(row, 3, float(itemlist[3]), fmt_money)
    worksheet.write_number(row, 4, float(itemlist[4]), fmt_money)
    # calculate the row and column for the formula
    worksheet.write_formula(row, 5, f"=E{row+1}-D{row+1}", fmt_money)

# add the conditional formatting
# Changed it from this:
#worksheet.conditional_format(1, 5, len(inventory_data), 5, {
    #"type": "cell",
    #"criteria": ">=",
    #"value": 0.75,
# To this:
# Change highlights the whole row when condition is met, not just the margin column
worksheet.conditional_format(1, 0, len(inventory_data), 5, {
    "type": "formula",
    "criteria" : "=$F2 >= .75",
    "format": fmt_cond
})

worksheet.set_zoom(150)
worksheet.autofit()

workbook.close()
'''


# Writing Workbook Properties:

# Using xlsx writer, you can programatically set the document properties of a workbook - 
#       Contain info about the workbook, such as title, author, descriptive words, and so on
# Can see these properties in Excel and going to properties viewer- 
#       Click on file, then click on info, click properties arrow, then click advanced properties
# While these properties don't directly affect the workbook, they can be read and used by external applications for a variety of reasons
#       Such as: custom workflows or search indexing

# XlsxWriter document properties

import xlsxwriter

workbook = xlsxwriter.Workbook("Properties.xlsx")
worksheet = workbook.add_worksheet()

# set the standard properties
# Takes a dictionary class with predefined values for the standard document properties
props = {
    "title": "Document Properties Example",
    "subject": "Shows how to use document properties in XlsxWriter",
    "author": "James Allen",
    "manager": "Colonel Monogram",
    "category": "Example Spreadsheets",
    "keywords": "Properties, Sample, XlsxWriter",
    "comments": "Created using XlsxWriter as a LinkedIn Learning Example"
}
workbook.set_properties(props)

# set some custom properties
# Can be used to store property values that are not within the standard set 
workbook.set_custom_property("Checked by", "Perry P")
workbook.set_custom_property("Approved", True)

workbook.close()