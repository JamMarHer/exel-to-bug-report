import sys
import os
import pyexcel_ods

def processOds(odsFile):
    pass


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print ('Please provide an ods file to process.')
        sys.exit()
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print ('File not found.')
        sys.exit()

    print ('Processing ods file: {}'.format(input_file))
    processOds(pyexcel_ods.get_data(input_file))
