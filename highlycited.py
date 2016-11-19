# Get xlrd with `pip install xlrd`
import xlrd

hcs2001 = xlrd.open_workbook("source/2001_HCR_as_of_September_8_2015.xlsx")
hcs2014 = xlrd.open_workbook("source/2001_HCR_as_of_September_8_2015.xlsx")
hcs2015 = xlrd.open_workbook("source/2001_HCR_as_of_September_8_2015.xlsx")
hcs2016 = xlrd.open_workbook("source/2001_HCR_as_of_September_8_2015.xlsx")

print("The number of worksheets is {0}".format(hcs2001.nsheets))
print("Worksheet name(s): {0}".format(hcs2001.sheet_names()))
sh = hcs2001.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
# rx in range(sh.nrows):
#    print(sh.row(rx))
