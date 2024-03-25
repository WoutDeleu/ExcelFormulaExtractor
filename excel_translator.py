from excel_helpers import *
from util import is_letter_ornumber

def get_sums(excel_formula):
    sums = []
    counter_closing_brackets_needed = 1
    current_sum = ''
    for ch in excel_formula:
        if ch == '(':
            counter_closing_brackets_needed += 1
        elif ch == ')':
            counter_closing_brackets_needed -= 1
            if counter_closing_brackets_needed == 0:
                sums.append(current_sum)
                current_sum = ''
        else:
            current_sum += ch
    return sums


def handle_sum(excel_formula, cells, formula):
    arguments_unformatted = excel_formula.split(',')
    for arg in arguments_unformatted:
        if is_sum(arg[:3]):
            arg = arg[4:]
            sums = get_sums(arg)
            for sum in sums:
                cells, formula = handle_sum(sum, cells, formula)

        elif ':' in arg:
            start_col, start_row = extract_col_row_from_excel_cell(arg.split(':')[0])
            end_col, end_row = extract_col_row_from_excel_cell(arg.split(':')[1])
            start_col_last_char = start_col[-1]
            start_col_except_last_char = start_col[:-1]
            end_col_last_char = end_col[-1]
            
            for i in range(int(start_row), int(end_row)+1):
                for j in range(ord(start_col_last_char), ord(end_col_last_char) + 1):
                    cells.append(Cell('Tax Calculation', start_col_except_last_char + chr(j) + str(i)))
                    formula += start_col_except_last_char + chr(j) + str(i) + '+'

        else:
            cells.append(Cell('Tax Calculation', arg))
            formula += arg + '+'
    
    return cells, formula[:-1]


def  extract_formula_cells(excel_formula):
    # TODO handle references to other sheets - nu, default sheet is 'Tax Calculation'
    
    cells = []
    formula = ''
    
    # remove first character (=)
    excel_formula = remove_char_from_string(excel_formula, 0)
    
    # TODO split up excel formula in operators and parts
    operators, parts = split_string_operators(excel_formula)
    
    # handle first 
    if excel_formula[0] == '-' or excel_formula[0] == '+':
        excel_formula = remove_char_from_string(excel_formula, 0)
    
    if is_sum(excel_formula[:3]):
        excel_formula = excel_formula[4:]
        
        sums = get_sums(excel_formula)
        for sum in sums:
            cells, formula = handle_sum(sum, cells, formula)
        
    # TODO Vergeet de AND and OR niet!
    elif is_if(excel_formula[:2]):
        pass
    elif is_iferror(excel_formula[:7]):
        pass
    elif is_excel_cell(excel_formula):
        pass
    elif excel_formula[0] == '-':
        pass
    else:
        print(excel_formula)
        raise Exception('Invalid formula')
    
    return cells, formula




def split_string_operators(string):
    # Queues
    operators = []
    parts = []
    
    current_part = ''    
    brackets_to_close = 0
    
    building_up_a_part = False
    is_allowed_to_close = False
    
    
    i = 0
    while i < len(string):
        if brackets_to_close > 0:
            current_part += string[i]
            if string[i] == ')':
                brackets_to_close -= 1
            if string[i] == '(':
                brackets_to_close += 1
            
        elif is_sum(string[i:i+3]):
            current_part += 'SUM('
            brackets_to_close += 1
            i += 3
        
        elif is_if(string[i:i+2]):
            current_part += 'SUM('
            brackets_to_close += 1
            i += 2
        
        elif string[i] == '(':
            current_part += string[i]
            brackets_to_close += 1
        
        elif is_operator(string[i]) and is_allowed_to_close:
            operators.append(string[i])
            parts.append(current_part)
            
        else:
            current_part += string[i]
            if is_letter_or_number(string[i]):
                is_allowed_to_close = True
        i += 1
    return operators, parts


# Voorbeeld
input_string = "-A + SUM(A) + B + C"
operators, parts = split_string_operators(input_string)
print("Operatoren:", operators)
print("Onderdelen:", parts)
input_string = "-A + (-B) + C"
operators, parts = split_string_operators(input_string)
print("Operatoren:", operators)
print("Onderdelen:", parts)


# test the function
(extract_formula_cells('=SUM(A1:B3)'))


