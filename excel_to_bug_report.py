#!/usr/bin/env python
import sys
import os
import xlrd
import datetime

kIndentLength = 2
kIndent = {
    'version': 0,
    'bug': 0,
    'build-arguments': 0,
    'fix_revision': 1,
    'dataset': 0,
    'program': 0,
    'dockerfile': 0,
    'files': 0,
    'extra': 0,
    'system': 1,
    'bug-kind': 0,
    'url': 0,
    'description': 0,
}

VERSION = '0'
DATASET = 'robots'
PROGRAM = 'ardupilot'
DOCKERFILE = 'Dockerfile.bug'

def indent(s, size):
    indent = size * ' '
    return ''.join(indent+line for line in s.splitlines(True))


def processExcel(workbookInput):
    sheet = workbookInput.sheet_by_index(0)
    numColumns = sheet.ncols
    numRows = sheet.nrows
    print('Columns: {}; Rows: {}'.format(numColumns, numRows))

    bugs = []
    for row in range(1, numRows):
        bug = {
            'version': VERSION,
            'bug': sheet.cell_value(rowx=row, colx=6),
            'build-arguments': '',
            'fix_revision': sheet.cell_value(rowx=row, colx=6),
            'dataset': DATASET,
            'program': PROGRAM,
            'dockerfile': DOCKERFILE,
            'files': sheet.cell_value(rowx=row, colx=3),
            'extra': '',
            'system': sheet.cell_value(rowx=row, colx=4),
            'bug-kind': sheet.cell_value(rowx=row, colx=5),
            'url': sheet.cell_value(rowx=row, colx=2),
            'description': sheet.cell_value(rowx=row, colx=1)
        }
        bugs.append(bug)

    reportBugs(bugs)


def reportBugs(bugs):
    with open('template', 'r') as f:
        template = f.read()

    for bug in bugs:
        if bug['bug'] == 'OK':
            continue

        report = template
        for (k, v) in bug.items():
            if k == 'files':
                v = '\n'.join(['- {}'.format(f) for f in v.split('\n')])

            # indent
            if kIndent[k] > 0:
                v = indent(v, kIndent[k] * kIndentLength)

            report = report.replace('__{}__'.format(k.upper()), v)
        # write to file
        with open('bugs/{}.bug.yaml'.format(bug['bug']), 'w') as f:
            f.write(report)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Please provide a .xlsx file to process.')
        sys.exit()
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print('File not found.')
        sys.exit()

    print('Processing xlsx file: {}'.format(input_file))
    processExcel(xlrd.open_workbook(input_file))
