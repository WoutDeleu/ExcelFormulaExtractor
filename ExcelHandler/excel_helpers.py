import traceback
import openpyxl
import re
from Util.util import *
from Util.Cell import *
from openpyxl.worksheet.formula import ArrayFormula
import datetime

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

def is_fully_covered_by_brackets(string):
    if string[0] != '(':
        return False
    count_opened_brackets = 0
    for ch in string[:-1]:
        if ch == '(':
            count_opened_brackets += 1
        if ch == ')':
            count_opened_brackets -= 1
        if count_opened_brackets == 0:
            return False
    return True


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

def is_round(string):
    return string == 'ROUND'

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

def is_exception(string):
    return  string == "/" or type(string) == ArrayFormula or isinstance(string, datetime.time) or string == None or string == '' or string == 'Ottignies- Louvain- La-Neuve'


def is_excel_range(string):
    # # Pattern to check for Excel range format
    pattern = r'^[A-Z]+[1-9]\d*:[A-Z]+[1-9]\d*$'
    pattern_absolute = r'^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$'
    strings = string.split('!')
    if len(strings) == 1:
        return bool(re.match(pattern, string)) or bool(re.match(pattern_absolute, string))
    elif len(strings) == 2:
        if strings[0][0] == '\'':
            return bool(re.match(pattern, strings[1])) or bool(re.match(pattern_absolute, strings[1]))
        else:
            checking_string = strings[0].split(':')[0] + ':' + strings[1]
            return bool(re.match(pattern, checking_string)) or bool(re.match(pattern_absolute, checking_string))
    elif len(strings) == 3:
        checking_string = strings[1].split(':')[0] + ':' + strings[2]
        return bool(re.match(pattern, checking_string)) or bool(re.match(pattern_absolute, checking_string))
    
def is_led_by_sheetname(string):
    return string[0] == '\'' or string.replace(' ', '')[0] == '\''

def handle_leading_space(string):
    if string[0] == ' ':
        return handle_leading_space(string[1:])
    return string

def is_or(string):
    return string == 'OR'

def is_and(string):
    return string == 'AND'

def is_not(string):
    return string == 'NOT'

def is_xor(string):
    return string == 'XOR'

# Absolute and relative cell references
def is_absolute_reference(cell_reference):
    """
    Check if a cell reference is absolute or not.
    
    Parameters:
        cell_reference (str): The cell reference to check.
        
    Returns:
        bool: True if the cell reference is absolute, False otherwise.
    """
    # Absolute references in Excel start with a "$" sign
    if "$" in cell_reference:
        return True
    else:
        return False

def absolute_to_relative(cell_reference):
    """
    Convert an absolute cell reference to a relative cell reference.
    
    Parameters:
        cell_reference (str): The absolute cell reference to convert.
        
    Returns:
        str: The relative cell reference.
    """
    # Als het een absolute referentie is, verwijder dan alle "$" tekens
    if "$" in cell_reference:
        cell_reference = cell_reference.replace("$", "")
        
    return cell_reference


def handle_absolute_relative_cell(cell):
    if is_absolute_reference(cell) and is_excel_cell(absolute_to_relative(cell)):
        return absolute_to_relative(cell)
    else:
        return cell