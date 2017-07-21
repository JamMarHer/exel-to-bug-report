#!/usr/bin/env python
import sys
import os
import xlrd
import json

def processExcel(workbookInput):
    sheet = workbookInput.sheet_by_index(0)
    numColumns = sheet.ncols
    numRows   = sheet.nrows
    print('Columns: {}; Rows: {}'.format(numColumns, numRows))
    formattedData = {}
    for row in range(numRows):
        commit = sheet.cell_value(rowx=row, colx=0)
        formattedData[commit] = []
        for col in range(1, numColumns):
            formattedData[commit].append(sheet.cell_value(rowx=row, colx=col))
    reportBugs(formattedData)


def reportBugs(formattedData):
    with open ('format', 'r') as f:
        # why _format?
        _format = [l.strip() for l in f]

    for commit in formattedData:
        count = 0
        # CT: what is [21][0:6]? The first 6 chars of the SHA?
        # If so, why not commit[:6]?
        with open('bugs/{}.bug'.format(formattedData[commit][21][0:6]), 'w') as bugReport:
            for line in _format:
                if 'bug:' in line:
                    bugReport.write('bug: \n')
                    continue
                if 'fix:' in line:
                    bugReport.write('fix: \n')
                    continue

                lineToWrite += str(' {}'.format(formattedData[commit][count]).replace('\n',''))
                bugReport.write(lineToWrite)
                bugReport.write('\n\n')
                count += 1

            # why? this is dead code
            count = 0

    print('Done.')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        # CT: are we processing .ods or .xlsx?
        print('Please provide an ods file to process.')
        sys.exit()
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print('File not found.')
        sys.exit()

    print('Processing xlsx file: {}'.format(input_file))
    processExcel(xlrd.open_workbook(input_file))
