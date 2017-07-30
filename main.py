import os
from pandas import ExcelFile, read_excel
from re import sub
import sys
from inspect import getsourcefile
from os import path
import argparse
from Excel import *


def List_Files(path):
    for roots, dirs, files in os.walk(path, topdown=True):
        for name in files:
            if name.lower().endswith(('.xls', '.xlsx')):
                Excel_file(os.path.join(roots, name), name, target)
        for name in dirs:
            if name.lower().endswith(('.xls', '.xlsx')):
                Excel_file(os.path.join(roots, name), name, target)


def Create_File(file):
    global target
    target = open(file, 'w+')
    target.write('''SET SAFETY OFF
SET VERIFY OFF
CLOSE
CLOSE SECONDARY

SET FOLDER /_01_Import

EXECUTE 'cmd /c MD ".\FIL"'
    ''')


def Close_file():
    target.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ACL Analytics' import script \
    generator for Excel and CSV files")
    parser.add_argument("-d", "--directory", help="Path to directory where files \
    to be imported are located", default=os.path.split(
                                    os.path.abspath(__file__))[0]+"\\")
    parser.add_argument("-o", "--output", help="Path or name where the import \
    script will be saved", required=True)
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)
    args = parser.parse_args()
    Create_File(args.output)
    List_Files(args.directory)
    Close_file()
