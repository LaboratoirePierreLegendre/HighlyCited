# Get xlrd with `pip install xlrd`
import xlrd
import pprint

# Utility functions

def create_set(sheet, categoryColum):
    # Given a sheet, and the column in which the category is found, 
    # we will create a string for each row and add it to the set.
    # A set is just an unsorted bag of unique values, so if the same 
    # string is added twice, only one will remain in the set

    "Fill out a set with the family name, first+middle name, and category (ignore affiliation for now)"
    resultSet = set()
    # Note: we skip the first line so we start at 1 (range is 0-based)
    for rx in range(1, sheet.nrows):
        firstname = sheet.row(rx)[0].value # The firstname should be column 0 (or 'A') of the sheet
        lastname = sheet.row(rx)[1].value # The lastname should be column 1 (or 'B') of the sheet
        category = sheet.row(rx)[categoryColum].value # The category column is specified as a parameter when you call this function
        individual = lastname.strip() + "," + firstname.strip() + "," + category.strip() # We use strip() to remove whitespace at the beginning or end of each string
        resultSet.add(individual)
    return resultSet

# Program starts here

print "Loading all HCR files..."

hcr2001 = xlrd.open_workbook("source/2001_HCR_as_of_September_8_2015.xlsx")
hcr2014 = xlrd.open_workbook("source/2014_HCR_as_of_September_8_2015.xlsx")
hcr2015 = xlrd.open_workbook("source/2015_HCR_as_of_December_1_2015.xlsx")
hcr2016 = xlrd.open_workbook("source/2016_HCR_as_of_November_16_2016.xlsx")

# Un-comment this if you want to see some stats about the sheets
# print("Some statistics about the sheets we just loaded:")
# print("hcs2001 has {0} worksheets".format(hcs2001.nsheets))
sh2001 = hcr2001.sheet_by_index(0)
# print("hcs2014 has {0} worksheets".format(hcs2014.nsheets))
sh2014 = hcr2014.sheet_by_index(0)
# print("hcs2015 has {0} worksheets".format(hcs2015.nsheets))
sh2015 = hcr2015.sheet_by_index(0)
# print("hcs2016 has {0} worksheets".format(hcs2016.nsheets))
sh2016 = hcr2016.sheet_by_index(0)

# print("sh2001 ({0}) has {1} rows and {2} columns".format(sh2001.name, sh2001.nrows, sh2001.ncols))
# print("sh2001 Cell A1 is {0}".format(sh2001.cell_value(rowx=0, colx=0)))
#
# print("sh2014 ({0}) has {1} rows and {2} columns".format(sh2014.name, sh2014.nrows, sh2014.ncols))
# print("sh2014 Cell A1 is {0}".format(sh2014.cell_value(rowx=0, colx=0)))
#
# print("sh2015 ({0}) has {1} rows and {2} columns".format(sh2015.name, sh2015.nrows, sh2015.ncols))
# print("sh2015 Cell A1 is {0}".format(sh2015.cell_value(rowx=0, colx=0)))
#
# print("sh2016 ({0}) has {1} rows and {2} columns".format(sh2016.name, sh2016.nrows, sh2016.ncols))
# print("sh2016 Cell A1 is {0}".format(sh2016.cell_value(rowx=0, colx=0)))

# Create the sets for each year
# Note for columns: column 'A' is column 0, 'B' is 1, etc...
set2001 = create_set(sh2001, 3) # In 2001, column 'D' contains the categories
set2014 = create_set(sh2014, 2) # In 2014, column 'C' contains the categories
set2015 = create_set(sh2015, 2) # In 2015, column 'C' contains the categories
set2016 = create_set(sh2016, 2) # In 2016, column 'C' contains the categories

# We now have four sets containing the HCR data for each year
# You can print each set individually, or use set operators to check if there is
# an intersection between them, or print the union of both sets

print ("\n")
print ("Result set")
# Use pprint now to pretty-print the set
# I like to sort the results, so I put everything in a sorted() call
# Modify at will...
pprint.pprint( sorted( set2016.intersection( set2015 ).intersection( set2014 ).intersection( set2001 ) ) )
