# Get xlrd with `pip install xlrd`
import xlrd
import pprint

def create_set(sheet):
    "Fill out a set with the family name, first+middle name, and category (ignore affiliation for now)"
    resultSet = set()
    # Note: we skip the first line so we start at 1 (range is 0-based)
    for rx in range(1, sheet.nrows):
        firstname = sheet.row(rx)[0].value
        lastname = sheet.row(rx)[1].value
        affiliation = sheet.row(rx)[2].value
        individual = lastname + "," + firstname + "," + affiliation
        resultSet.add(individual)
    return resultSet

print "Loading all HCR files..."

hcs2001 = xlrd.open_workbook("source/2001_HCR_as_of_September_8_2015.xlsx")
hcs2014 = xlrd.open_workbook("source/2014_HCR_as_of_September_8_2015.xlsx")
hcs2015 = xlrd.open_workbook("source/2015_HCR_as_of_December_1_2015.xlsx")
hcs2016 = xlrd.open_workbook("source/2016_HCR_as_of_November_16_2016.xlsx")

print("Some statistics about the sheets we just loaded:")
print("hcs2001 has {0} worksheets".format(hcs2001.nsheets))
sh2001 = hcs2001.sheet_by_index(0)
print("hcs2014 has {0} worksheets".format(hcs2014.nsheets))
sh2014 = hcs2014.sheet_by_index(0)
print("hcs2015 has {0} worksheets".format(hcs2015.nsheets))
sh2015 = hcs2015.sheet_by_index(0)
print("hcs2016 has {0} worksheets".format(hcs2016.nsheets))
sh2016 = hcs2016.sheet_by_index(0)

print("sh2001 ({0}) has {1} rows and {2} columns".format(sh2001.name, sh2001.nrows, sh2001.ncols))
print("sh2001 Cell A1 is {0}".format(sh2001.cell_value(rowx=0, colx=0)))

print("sh2014 ({0}) has {1} rows and {2} columns".format(sh2014.name, sh2014.nrows, sh2014.ncols))
print("sh2014 Cell A1 is {0}".format(sh2014.cell_value(rowx=0, colx=0)))

print("sh2015 ({0}) has {1} rows and {2} columns".format(sh2015.name, sh2015.nrows, sh2015.ncols))
print("sh2015 Cell A1 is {0}".format(sh2015.cell_value(rowx=0, colx=0)))

print("sh2016 ({0}) has {1} rows and {2} columns".format(sh2016.name, sh2016.nrows, sh2016.ncols))
print("sh2016 Cell A1 is {0}".format(sh2016.cell_value(rowx=0, colx=0)))

# Create our empty sets
set2001 = create_set(sh2001)
set2014 = create_set(sh2014)
set2015 = create_set(sh2015)
set2016 = create_set(sh2016)

print ("\n")
print ("Result set")
# Use pprint now to pretty-print the set
# I like to sort the results, so I put everything in a sorted() call
# Modify at will...
pprint.pprint( sorted( set2016.intersection( set2015 ).intersection( set2014 ) ) )
