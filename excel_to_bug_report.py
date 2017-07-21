#!/usr/bin/env python
import sys
import os
import xlrd
import json

kIndentLength = 2
kIndent = {
    'title': 0,
    'description': 1,
    'classification': 0,
    'keywords': 0,
    'severity': 0,
    'links': 1,
    'phase': 0,
    'specificity': 0,
    'languages': 0,
    'detected_by': 0,
    'reported_by': 0,
    'issue': 0,
    'time_reported': 0,
    'hash': 0,
    'pull_request': 0,
    'files': 2,
    'time_fixed': 0
}


def indent(s, size):
    indent = size * ' '
    return ''.join(indent+line for line in s.splitlines(True))


def processExcel(workbookInput):
    sheet = workbookInput.sheet_by_index(0)
    numColumns = sheet.ncols
    numRows = sheet.nrows
    print('Columns: {}; Rows: {}'.format(numColumns, numRows))

    bugs = []
    for row in range(numRows):
        bug = {
            'title': sheet.cell_value(rowx=row, colx=1),
            'description': sheet.cell_value(rowx=row, colx=2),
            'classification': sheet.cell_value(rowx=row, colx=3),
            'keywords': sheet.cell_value(rowx=row, colx=4),
            'severity': sheet.cell_value(rowx=row, colx=5),
            'links': sheet.cell_value(rowx=row, colx=6),
            'phase': sheet.cell_value(rowx=row, colx=7),
            'specificity': sheet.cell_value(rowx=row, colx=8),
            'languages': sheet.cell_value(rowx=row, colx=9).split('\n'),
            'detected_by': sheet.cell_value(rowx=row, colx=10),
            'reported_by': sheet.cell_value(rowx=row, colx=11),
            'issue': sheet.cell_value(rowx=row, colx=12),
            'time_reported': sheet.cell_value(rowx=row, colx=13),
            'hash': sheet.cell_value(rowx=row, colx=14),
            'pull_request': sheet.cell_value(rowx=row, colx=15),
            'files': sheet.cell_value(rowx=row, colx=16).split('\n'),
            'time_fixed': sheet.cell_value(rowx=row, colx=17)
        }
        if bug['title'] != '':
            bugs.append(bug)
    reportBugs(bugs)


def reportBugs(bugs):
    with open('template', 'r') as f:
        template = f.read()

    for bug in bugs:
        report = template
        for (k, v) in bug.items():
            if k == 'languages':
                v = ','.join(v)
            elif k == 'files':
                v = '\n'.join(v)

            # indent
            if kIndent[k] > 0:
                v = indent(v, kIndent[k] * kIndentLength)

            report = report.replace('__{}__'.format(k.upper()), v)

        print(report)


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
