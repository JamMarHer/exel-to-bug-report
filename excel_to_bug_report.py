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

    # CT: why are we storing in a hash map based on the commit link?
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
        _format = [l.strip('\n') for l in f]

    for commit in formattedData:
        count = 0
        short_hash = formattedData[commit][21][0:6]
        print(commit)
        with open('bugs/{}.bug'.format(short_hash), 'w') as bugReport:
            for line in _format:
                if 'bug:' in line:
                    bugReport.write('bug:\n')
                    continue
                if 'fix:' in line:
                    bugReport.write('fix:\n')
                    continue

                # CT: why lineToWrite and line?
                line = '{} {}\n\n'.format(line, formattedData[commit][count])
                bugReport.write(line)
                count += 1 # use enumerate?

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
