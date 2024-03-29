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

def split_up_formulas(excel_formula):
    current_part = ''
    parts = []
    brackets_to_close = 0
    for ch in excel_formula:
        if ch == '(':
            brackets_to_close += 1
            current_part += ch
        elif ch == ')':
            brackets_to_close -= 1  
            current_part += ch
        elif (ch == ',' or ch == ';') and brackets_to_close == 0:
            parts.append(current_part)
            current_part = ''
        else:
            current_part += ch
    parts.append(current_part)
    return parts

def split_up_conditions(condition):
    # Lijst met operatoren en hun bijbehorende patronen
    operators = {
        '>': '>',
        '<': '<',
        '=': '=',
        '>=': '>=',
        '<=': '<=',
        '<>': '<>'
    }
    
    parts = []
    used_operators = []
    current_part = ''
    i=0
    while i < len(condition):
        if condition[i:i+2] in operators:
            parts.append(current_part)
            if condition[i:i+2] == '<>':
                used_operators.append('!=')
            else: 
                used_operators.append(condition[i:i+2])
            current_part = ''
            i += 1
        elif condition[i] in operators:
            parts.append(current_part)
            if condition[i] == '=':
                used_operators.append('==')
            else:
                used_operators.append(condition[i])
            current_part = ''
        else:
            current_part += condition[i]
        i += 1
    parts.append(current_part)
    return parts, used_operators


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

def is_VLOOKUP(string):
    return string == 'VLOOKUP'

def is_excel_cell(string):
    pattern = r'^[A-Z]+[0-9]+$'
    return bool(re.match(pattern, string))

def is_number(string):
    patroon = r'^[-+]?\d*\.?\d+(,\d+)?$'
    return bool(re.match(patroon, string))

def is_operator(char):
    if char in ['+', '-', '*', '/']:
        return True
    return False

def is_excel_range(string):
    # Pattern to check for Excel range format
    pattern = r'^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$'
    return bool(re.match(pattern, string))

def is_or(string):
    return string == 'OR'

def is_and(string):
    return string == 'AND'

def is_not(string):
    return string == 'NOT'

def is_xor(string):
    return string == 'XOR'