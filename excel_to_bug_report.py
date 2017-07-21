#!/usr/bin/env python
import sys
import os
import xlrd
import json

def processExcel(workbookInput):
    sheet = workbookInput.sheet_by_index(0)
    numColumns = sheet.ncols
    numRows   = sheet.nrows
    print ('Columns: {}; Rows: {}'.format(numColumns, numRows))
    formatedData = {}
    for row in range(numRows):
        commit = sheet.cell_value(rowx=row, colx=0)
        formatedData[commit] = []
        for col in range(1, numColumns):
            formatedData[commit].append(sheet.cell_value(rowx=row, colx=col))
    reportBugs(formatedData)


def reportBugs(formatedData):
    with open ('format', 'r') as input_format:
        _format = input_format.readlines()
    for commit in formatedData:
        count = 0
        with open('bugs/{}.bug'.format(formatedData[commit][21][0:6]), 'a') as bugReport:
            for line in _format:
                if 'bug:' in line:
                    bugReport.write('bug: \n')
                    continue
                if 'fix:' in line:
                    bugReport.write('fix: \n')
                    continue
                lineToWrite = line.replace('\n','') 
                lineToWrite += str(' {}'.format(formatedData[commit][count]).replace('\n',''))
                bugReport.write(lineToWrite)
                bugReport.write('\n\n')
                count += 1
            count = 0
    print ('Done.')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print ('Please provide an ods file to process.')
        sys.exit()
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print ('File not found.')
        sys.exit()

    print ('Processing xlsx file: {}'.format(input_file))
    processExcel(xlrd.open_workbook(input_file))
