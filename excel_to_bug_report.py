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
        print ('Commit: {}'.format(commit))
        formatedData[commit] = []
        for col in range(1, numColumns):
            formatedData[commit].append(sheet.cell_value(rowx=row, colx=col))
    for commit in formatedData:
        print (commit)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print ('Please provide an ods file to process.')
        sys.exit()
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print ('File not found.')
        sys.exit()

    print ('Processing ods file: {}'.format(input_file))
    processOds(xlrd.open_workbook(input_file))
