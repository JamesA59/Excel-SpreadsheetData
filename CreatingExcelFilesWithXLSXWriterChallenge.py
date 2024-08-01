# Creating Excel Files With XLSX Writer Challenge:

# Read a single CSV file and split it into multiple worksheets

import csv
import xlsxwriter


filename = "Inventory.csv"

# PUT YOUR CHALLENGE CODE HERE

def read_csv_to_array(filename):
    # define the array that will hold the data
    data = []
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data.append(row)
    return data

inventory_data = read_csv_to_array(filename)

workbook = xlsxwriter.Workbook("Inventory2.xlsx")
worksheet = workbook.add_worksheet("Inventory")

fmt_bold = workbook.add_format({'bold': True})
fmt_money = workbook.add_format(
    {'font_color': 'green', 'num_format': '$#,##0.00'})
fmt_cond = workbook.add_format({"bg_color": "#AAFFAA", "bold": True})

worksheet.write_row(0, 0, inventory_data[0], fmt_bold)
worksheet.write(0, 5, "Margin", fmt_bold)

for row, itemlist in enumerate(inventory_data[1:], start=1):
    worksheet.write(row, 0, itemlist[0])
    worksheet.write(row, 1, itemlist[1], fmt_bold)
    worksheet.write_number(row, 2, int(itemlist[2]))
    worksheet.write_number(row, 3, float(itemlist[3]), fmt_money)
    worksheet.write_number(row, 4, float(itemlist[4]), fmt_money)
    worksheet.write_formula(row, 5, f"=E{row+1}-D{row+1}", fmt_money)

worksheet.autofit()

workbook.close()
