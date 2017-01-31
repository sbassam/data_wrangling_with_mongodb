# -*- coding: utf-8 -*-
'''
Find the time and value of max load for each of the regions
COAST, EAST, FAR_WEST, NORTH, NORTH_C, SOUTHERN, SOUTH_C, WEST
and write the result out in a csv file, using pipe character | as the delimiter.

An example output can be seen in the "example.csv" file.
'''

import xlrd
import os
import csv
from zipfile import ZipFile

datafile = "2013_ERCOT_Hourly_Load_Data.xls"
outfile = "2013_Max_Loads.csv"


def open_zip(datafile):
    with ZipFile('{0}.zip'.format(datafile), 'r') as myzip:
        myzip.extractall()


def parse_file(datafile):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(0)
    data = []
    print sheet.nrows
    print sheet.ncols
    print sheet.cell_type (2, 4)
    for i in range(sheet.nrows):
        row = []
        for j in range(sheet.ncols):
            if sheet.cell_type(i, j) == 2:
                xl_date = sheet.cell_value(i,j)
                date = xlrd.xldate_as_tuple(xl_date, 0)
                row.append(date)
            else:
                row.append(sheet.cell_value(i,j))
        data.append(row)
    print data
    # YOUR CODE HERE
    # Remember that you can use xlrd.xldate_as_tuple(sometime, 0) to convert
    # Excel date to Python tuple of (year, month, day, hour, minute, second)
    return data

parse_file(datafile)
