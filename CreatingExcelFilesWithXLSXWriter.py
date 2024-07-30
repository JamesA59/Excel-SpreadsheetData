# Creating Excel Files With XLSX Writer:


# Introduction to Xlsx Writer:

# Unlike Openpyxl, this package is only used to write Excel spreadsheets-
#       You  can't read or manipulate spreadsheets
# Will learn how to use xlsx writer to create workbooks, style their contents, apply conditional formatting and filters, and create Excel tables


# Creating a workbook:


# Will create a workbook and add content to it

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
date_time = datetime.datetime.strptime('2030-07-28', '%y-%m-%d')
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