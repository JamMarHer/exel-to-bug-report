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



# Done this way since we don't have the sheet in correct format. Can be easily
# improved after we have it nicely formated.
def reportBugs(formatedData):
    for commit in formatedData:
        with open('bugs/{}.bug'.format(formatedData[commit][11]), 'a') as bugReport:
            bugReport.write('title: {}\n\n'.format(formatedData[commit][0]))
            bugReport.write('description: {}\n\n'.format(formatedData[commit][1]))
            bugReport.write('classification: {}\n\n'.format(formatedData[commit][2]))
            bugReport.write('keywords: {}\n\n'.format(formatedData[commit][3]))
            bugReport.write('severity{}\n\n'.format(formatedData[commit][4]))
            bugReport.write('links{}\n\n'.format(formatedData[commit][17]))
            bugReport.write('bug: \n\n')
            bugReport.write('  phase: {}\n\n'.format(formatedData[commit][5]))
            bugReport.write('  specifity: {}\n\n'.format(formatedData[commit][6]))
            bugReport.write('  architectural-location: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  application{}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  task: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  subsystem: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  package: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  language: {}\n\n'.format(formatedData[commit][14]))
            bugReport.write('  detected-by: {}\n\n'.format(formatedData[commit][7]))
            bugReport.write('  reported-by: {}\n\n'.format(formatedData[commit][8]))
            bugReport.write('  isssue: {}\n\n'.format(formatedData[commit][9]))
            bugReport.write('  time-reported: {}\n\n'.format(formatedData[commit][10]))
            bugReport.write('  reproducibility: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  trace: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('fix: \n\n')
            bugReport.write('  repo: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  hash: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  pull-request: {}\n\n'.format(formatedData[commit][12]))
            bugReport.write('  license: {}\n\n'.format(formatedData[commit][11]))
            bugReport.write('  fix-in: {}\n\n'.format(formatedData[commit][13]))
            bugReport.write('  language: {}\n\n'.format(formatedData[commit][14]))
            bugReport.write('  time: {}\n\n'.format(formatedData[commit][15]))
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
