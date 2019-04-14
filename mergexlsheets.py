"""Merges sheets of an excel file along a given primary key with NAN replaced with a custom character and saves the
output to a csv file"""
# this is a class file

import os
import pandas as pd
from functools import reduce

class MergeXLsheets():

    def __init__(self, XLfilename, mergeid, fillnachar, how='outer'):

        self.XLfilename = XLfilename

        if not os.path.isfile(XLfilename):
            print("ERROR: File not founc." + XLfilename)
            exit(-1)
        else:
            self.xls = pd.ExcelFile(XLfilename)

        # database to store merged databases from different XL sheets
        self.db = self.mergedatabases(self.xls, mergeid, fillnachar, how)

    def printdatabases(self):
        """Prints all databases from all sheets"""
        print(self.db)

    def printsheetnames(self):
        """Prints the sheet names only"""
        print(self.xls.sheet_names)

    # TODO: Decide on the how="outer" vs how="col" options

    def mergedatabases(self, xls, mergeid, fillnachar, how="col"):

        if how=="col":
            mergedatabasescolumnwise(self, xls, mergeid, fillnachar, "outer")
        else:
            mergedatabasesrowwise(self, xls, mergeid, fillnachar, "outer")

    # TODO: add another function to merge sheets vertically: mergedatabasesrowwise()

    # TODO: Add function to export every sheet to their own CSV file


    def mergedatabasescolumnwise(self, xls, mergeid, fillnachar, how):
        """Merges XL sheets"""

        data_frames = []

        print("Extracting databases ...")

        for sheet in xls.sheet_names:
            # db[i] = xls.parse(sheet)
            # same as
            df = pd.read_excel(xls, sheet)
            data_frames.append(df)

        print("Merging databases ...")

        # merge multiple sheets by selecting two at a time with lambda function
        # Reduce applies a rolling computation to sequential pairs of values in a list. Good for factorials.
        # duplicates columns get(_x,_y) suffixes by default. VChnage to (,_y) so that we can remove _y columns
        df_merged = reduce(lambda left, right: pd.merge(left, right, on=mergeid, how=how, suffixes=('', '_y')).fillna(fillnachar), data_frames)

        # drop duplicate columns with _y suffix
        to_drop = [x for x in df_merged if x.endswith('_y')]
        df_merged.drop(to_drop, axis=1, inplace=True)

        df_merged.set_index(mergeid, inplace=True)

        return df_merged

    def savetocsv(self, csvname='dbmerged.csv'):
        """Saves merged database to CSV file"""
        self.db.to_csv(csvname)


