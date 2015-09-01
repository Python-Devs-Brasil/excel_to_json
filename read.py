#!/usr/bin/python
# -*- coding: utf-8 -*-

# See full documentation about xlrd module https://secure.simplistix.co.uk/svn/xlrd/trunk/xlrd/doc/xlrd.html?p=4966
import xlrd
# Import json module to export data
import json
# Import dictionary module
import dictionary
# Import sys module
import sys


reload(sys)
sys.setdefaultencoding("utf-8")

jsonObject = []

def getValues(objectXls, dictionary):
    global jsonObject
    book = objectXls
    recebe = []
    sheetToWork = book.sheet_by_index(2)
    # Count total of lines of result
    countRows = sheetToWork.nrows

    i = 1
    response = {}
    while i < countRows:
        cells = sheetToWork.row_slice(rowx=i,
                                    start_colx=0,)

        for (value, key) in zip(cells, dictionary):
            response[str(key)] = '%s' % str(value.value)

        i += 1
        jsonObject.append(response.copy())
    pass

# Open xlsx file to open
book = xlrd.open_workbook("example.xlsx")
# set columns to dicitionary
dictionary.setDictionaryFields(book)
getValues(book, dictionary.getDictionaryFields())



with open("export.json", "w+") as file:
    json.dump(jsonObject , file)
