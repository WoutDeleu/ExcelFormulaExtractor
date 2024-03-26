from ExcelHandler.excel_helpers import *
from ExcelHandler.handle_max_min import handle_max_calculation
from ExcelHandler.handle_sum import handle_sum_calculation
from Util.util import is_letter_or_number
from Util.DataStructures import Queue


def add_to_resulting_formula(resulting_formula, formula, operators):
    resulting_formula += formula
    if not operators.is_empty():
        resulting_formula += operators.pop()
    return resulting_formula


def  extract_formula_cells(excel_formula):
    # TODO handle references to other sheets - nu, default sheet is 'Tax Calculation'
    
    cells = []
    formula = ''
    
    # remove first character (=)
    excel_formula = remove_char_from_string(excel_formula, 0)

    operators, parts = split_up_excel_formula(excel_formula)
    
    for element in parts.get_list():
        if is_sum(element[:3]):
            cells, current_formula = handle_sum_calculation(cells, element)
            formula = add_to_resulting_formula(formula, current_formula, operators)

        elif is_max(element[:3]):
            cells, current_formula = handle_max_calculation(cells, element)
            formula = add_to_resulting_formula(formula, current_formula, operators)
            
        elif is_min(element[:3]):
            pass
        
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



def split_up_excel_formula(string):
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
        # TODO make more elegant - integrate in previous while loop
        parts.add(current_part)
        
    return operators, parts
