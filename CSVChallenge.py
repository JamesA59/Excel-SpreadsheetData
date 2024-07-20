# CSV Challenge

import csv
import pprint
from decimal import Decimal

def read_csv_to_array(filename):
    data = []
    with open(filename, 'r') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            data.append(row)
    return data





def write_array_to_csv(data, filename):
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)


inventory_data = read_csv_to_array("Inventory.csv")
#pprint.pprint(inventory_data)

headers = inventory_data[0]
headers.append("Margin")

datarows = inventory_data[1:]
for row in datarows:
    margin_value = Decimal(row[4]) - Decimal(row[3])
    row.append(margin_value)

pprint.pprint(inventory_data)

write_array_to_csv(inventory_data, "CSVChallengeoutput.csv")