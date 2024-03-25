import traceback
import openpyxl
import re
from Util.util import *
from Util.Cell import *


def read_in_excel(path):
    workbook = openpyxl.load_workbook(filename=path)
    return workbook

def extract_col_row_from_excel_cell(cell):
    col = ''
    i = 0
    while not is_int(cell[i]):
        col += cell[i]
        i += 1
    row = cell[i:]
    return col, row


# Checks
def is_sum(string):
    return string == 'SUM'

def is_if(string):
    return string == 'IF'

def is_iferror(string):
    return string == 'IFERROR'

def is_max(string):
    return string == 'MAX'

def is_min(string):
    return string == 'MIN'

def is_excel_cell(string):
    pattern = r'^[A-Z]+[0-9]+$'
    return bool(re.match(pattern, string))

def is_operator(char):
    if char in ['+', '-', '*', '/']:
        return True
    return False

