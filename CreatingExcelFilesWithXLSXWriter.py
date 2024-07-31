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