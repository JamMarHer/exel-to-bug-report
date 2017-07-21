#!/usr/bin/env python
import sys
import os
import xlrd
import json
import datetime

kIndentLength = 2
kIndent = {
    'ready': 0,
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
    print(indent)
    return ''.join(indent+line for line in s.splitlines(True))


def processExcel(workbookInput):
    sheet = workbookInput.sheet_by_index(0)
    numColumns = sheet.ncols
    numRows = sheet.nrows
    print('Columns: {}; Rows: {}'.format(numColumns, numRows))

    bugs = []
    for row in range(1, numRows):
        bug = {
            'ready': sheet.cell_value(rowx=row, colx=0),
            'title': sheet.cell_value(rowx=row, colx=2),
            'description': sheet.cell_value(rowx=row, colx=3),
            'classification': sheet.cell_value(rowx=row, colx=4),
            'keywords': sheet.cell_value(rowx=row, colx=5),
            'severity': sheet.cell_value(rowx=row, colx=6),
            'links': sheet.cell_value(rowx=row, colx=7),
            'phase': sheet.cell_value(rowx=row, colx=8),
            'specificity': sheet.cell_value(rowx=row, colx=9),
            'detected_by': sheet.cell_value(rowx=row, colx=10),
            'reported_by': sheet.cell_value(rowx=row, colx=11),
            'issue': sheet.cell_value(rowx=row, colx=12),
            'time_reported': sheet.cell_value(rowx=row, colx=13),
            'hash': sheet.cell_value(rowx=row, colx=14),
            'pull_request': sheet.cell_value(rowx=row, colx=15),
            'files': sheet.cell_value(rowx=row, colx=16).split('\n'),
            'languages': sheet.cell_value(rowx=row, colx=17).split('\n'),
            'time_fixed': sheet.cell_value(rowx=row, colx=18)
        }
        bug['ready'] = bug['ready'] == 'Yes'

        if isinstance(bug['time_reported'], float):
            bug['time_reported'] = xlrd.xldate_as_tuple(bug['time_reported'], 0)
            bug['time_reported'] = datetime.datetime(*bug['time_reported'])
            bug['time_reported'] = bug['time_reported'].strftime('%d %B, %Y')

        if isinstance(bug['time_fixed'], float):
            bug['time_fixed'] = xlrd.xldate_as_tuple(bug['time_fixed'], 0)
            bug['time_fixed'] = datetime.datetime(*bug['time_fixed'])
            bug['time_fixed'] = bug['time_fixed'].strftime('%d %B, %Y')

        if bug['title'] != '':
            bugs.append(bug)

    reportBugs(bugs)


def reportBugs(bugs):
    with open('template', 'r') as f:
        template = f.read()

    for bug in bugs:
        if not bug['ready']:
            continue

        report = template
        for (k, v) in bug.items():
            if k == 'ready':
                continue
            if k == 'languages':
                v = ','.join(v)
            elif k == 'files':
                v = '\n'.join(['- {}'.format(f) for f in v])

            # indent
            if kIndent[k] > 0:
                v = indent(v, kIndent[k] * kIndentLength)

            report = report.replace('__{}__'.format(k.upper()), v)

        # write to file
        with open('bugs/{}.bug'.format(bug['hash'][:6]), 'w') as f:
            f.write(report)


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
