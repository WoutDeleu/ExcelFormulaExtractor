from excel_helpers import *
from Util.util import is_letter_or_number
from Util.DataStructures import Queue

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
    
    for element in parts.get_list():
        if is_sum(element[:3]):
            element = element[4:]
            
            sum_formula = ''
            sums = get_sums(element)
            for sum in sums:
                cells, sum_formula = handle_sum(sum, cells, sum_formula)
            formula += sum_formula
            if not operators.is_empty():
                formula += operators.pop()

        elif is_max(element[:3]):
            pass
        elif is_min(element[:3]):
            pass
        # TODO Vergeet de AND and OR niet!
        elif is_if(element[:2]):
            pass
        elif is_iferror(element[:7]):
            pass
        elif is_excel_cell(element):
            pass
        elif element[0] == '-':
            pass
        elif element[0] == '(':
            pass
        else:
            print(element)
            raise Exception('Invalid formula')
    
    return cells, formula




def split_string_operators(string):
    # Queues
    operators = Queue()
    parts = Queue()
    
    current_part = ''
    brackets_to_close = 0
    
    brackets_input_is_handled = True
    is_allowed_to_close = False
    
    i = 0
    while i < len(string):
        if brackets_to_close > 0:
            current_part += string[i]
            if string[i] == ')':
                brackets_to_close -= 1
            if string[i] == '(':
                brackets_to_close += 1
                
            if brackets_to_close == 0:
                brackets_input_is_handled = False
            
        elif is_sum(string[i:i+3]):
            current_part += 'SUM('
            brackets_to_close += 1
            i += 3
        
        elif is_if(string[i:i+2]):
            current_part += 'IF('
            brackets_to_close += 1
            i += 2
            
        elif is_iferror(string[i:i+7]):
            current_part += 'IFERROR('
            brackets_to_close += 1
            i += 7
            
        elif is_max(string[i:i+3]):
            current_part += 'MAX('
            brackets_to_close += 1
            i += 3
        
        elif is_min(string[i:i+3]):
            current_part += 'MIN('
            brackets_to_close += 1
            i += 3
        
        elif string[i] == '(':
            current_part += string[i]
            brackets_to_close += 1
        
        elif is_operator(string[i]) and is_allowed_to_close:
            operators.add(string[i])
            parts.add(current_part)
            current_part = ''
            brackets_input_is_handled = True
            
        else:
            current_part += string[i]
            if is_letter_or_number(string[i]):
                is_allowed_to_close = True
        i += 1
        
    if not brackets_input_is_handled:
        parts.add(current_part)
        
    return operators, parts

# test the function
(extract_formula_cells('=SUM(A1:B3)'))


