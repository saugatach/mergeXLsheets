#!/home/jones/anaconda3/bin/python3

import os
import sys
from mergexlsheets import MergeXLsheets

print("Usage: ./merge.py <XL file> [optional args .. <mergeoncolumnID = \"ID\"> <outputfile = \"dbmerged.csv\"> <NANchar = \"?\"> ] \n")

# set argv options
nargs = len(sys.argv)
if nargs == 1:
    print("Must provide input filename as argument. Exiting ....")
    exit(-1)
elif nargs == 2:
    XLfilename, mergeid, csvname, fillnachar = [sys.argv[1], "ID", "dbmerged.csv", '?']
elif nargs == 3:
    XLfilename, mergeid, csvname, fillnachar = [sys.argv[1], sys.argv[2], "dbmerged.csv", '?']
elif nargs == 4:
    XLfilename, mergeid, csvname, fillnachar = [sys.argv[1], sys.argv[2], sys.argv[3], '?']
else:
    XLfilename, mergeid, csvname, fillnachar = [sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]]

print("reading from:", XLfilename)
print("Merge ID:", mergeid)
print("Output filename:", csvname)
print("NAN character:", fillnachar)

# create new db from XL file and merge the sheets
newdb = MergeXLsheets(XLfilename, mergeid, fillnachar)
# newdb.printdatabases()
newdb.savetocsv(csvname)

print("Output saved to " + csvname)


