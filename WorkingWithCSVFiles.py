# Working with CSV files

# The CSV Format:

# Comma-Separated Values
# Data is stored in a flat format as rows of vlaues, separated by commas
# Widely used as common storage format for beneric spreadsheet data
# Useful for transfer of data between databases and spreadsheets
# Ex.
#       Ticker, Company, Open, Close
#       GOOG, Alphabet, 147.89, 149.21
#       META, Meta Platforms, 421.78, 456.23
#       MSFT, Microsoft Copr., 419.03, 403.75
#       TSLA, Tesla Inc., 173.84, 171.09
#       AMZN, Amazon.com Inc., 179.01, 175.24



# Reading CSV files into an array:

# Will read a CSV file into an array of arrays in Python
# Will need to use the CSV module in the Python Standard Library
# Downloading inventory.csv file and saved it to folder
# import the csv module from the standard library

''''
import csv

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

# Each row in the array is itself an array of values
print(f"Items: {len(inventory_data)}")
print(inventory_data[0])
print(inventory_data[1])
print(inventory_data[1][0], inventory_data[1][2]
'''


# Reading CSV files into a dictionary:

# A CSV file can be read as a dictionary as another option to an array of arrays
# Use the DictReader class to read a CSV file as a dictionary, which is also in the CSV module

'''
import csv
import pprint

def read_csv_to_dict(filename):
    data = {}
    with open(filename, 'r') as csvfile:
        reader = csv.DictReader(csvfile)

        #row = next(reader)
        # Below line prints the first row, a dictionary of key-value pairs
        #print(row)
        # Below prints the values that are used as keys in each row
        # These lines let you use item name as the overall dictionary key 
        #print(reader.fieldnames)

        # Will modify code to just use an integer for each key
        # Goes from this:
        #for row in reader:
            #data[row[reader.fieldnames[0]]] = row

        # To this:
        key = 0
        for row in reader:
            data[key] = row
            key +=1


    return data

# Example usage
inventory_data = read_csv_to_dict("Inventory.csv")

# Accessing data
# pprint is pretty print
pprint.pprint(inventory_data)

# Now that earlier code is changed to use integers as keys, this code is outdated
# Change this:
#pprint.pprint(inventory_data["Apple"])
#pprint.pprint(inventory_data["Apple"]["Consumer Price"])

# To this:
pprint.pprint(inventory_data[0])
pprint.pprint(inventory_data[0]["Consumer Price"])

# Each item is now differentiated by a number instead of it's name, but is still essentially the same
'''


# Reading CSV files with a filter:

# When working with a large CSV file, it'd be beneficial to just work with a subset of the overall data
# Can accomplish this by defining a filter function and apply it to each row as the data is read

'''
import csv
import pprint

def read_csv_filter_rows(filename, filter_func):
  # array to hold the filtered data result
  filtered_data = []

  with open(filename, 'r') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
      if (filter_func(row)):
        filtered_data.append(row)
  return filtered_data

# Filter function (replace with your specific filtering criteria)
def filter_by_catagory(row, category):
  return row[1] == category

# Call the read function with a filter function
# Should filter out everyting that isn't a fruit
filtered_rows = read_csv_filter_rows("Inventory.csv", lambda row: filter_by_catagory(row, "Fruits"))

# Print filtered data
pprint.pprint(filtered_rows)
'''


# Writing a CSV file:

# The CSV module provides a set of classes and functions for writing CSV files
# Will write data in an array format
'''
import csv

# Sample data
data = [
  ["Item Name", "Category", "Quantity", "Wholesale Price", "Consumer Price"],
  ["Apple","Fruits",100,0.50,0.75],
  ["Banana","Fruits",150,0.35,0.50],
  ["Orange","Fruits",120,0.45,0.65],
  ["Grapes","Fruits",80,0.60,0.85],
  ["Strawberries","Fruits",90,1.20,1.50]
]

def write_array_to_csv(data, filename):
    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(data)
# Write data to CSV file
# Creates an Excel sheet titled output.csv in same folder as code
write_array_to_csv(data, "output.csv")
'''

