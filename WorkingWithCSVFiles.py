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

import csv

def read_csv_to_array(filename):
  # define the array that will hold the data
  data = []

# Read the data into an array of arrays
inventory_data = read_csv_to_array("Inventory.csv")

# Each row in the array is itself an array of values