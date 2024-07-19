# CSV Challenge

import csv
import pprint

def read_csv_to_dict(filename):
    data = {}
    with open(filename, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        key = 0
        for row in reader:
            data[key] = row
            key +=1
    return data

inventory_data = read_csv_to_dict("Inventory.csv")
'''
# Accessing data
# pprint is pretty print
pprint.pprint(inventory_data)

# To this:
pprint.pprint(inventory_data[0])
pprint.pprint(inventory_data[0]["Consumer Price"])

# Each item is now differentiated by a number instead of it's name, but is still essentially the same
'''

data = [inventory_data[0]]
fieldnames= ["Item Name", "Category", "Quantity", "Wholesale Price", "Consumer Price", "Margin" ]

print(data)