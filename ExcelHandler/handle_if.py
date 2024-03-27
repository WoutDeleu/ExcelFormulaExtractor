import ExcelHandler
from Util.Cell import Cell

def split_up_if_formula(string):
    condition_parts = []
    
    brackets_to_close = 0
    current_part = ''
    
    i = 0
    while i < len(string):
        if string[i] == '(':
            brackets_to_close += 1
        if string[i] == ')':
            brackets_to_close -= 1
            
        
        if brackets_to_close == 0 and (string[i] == ',' or string[i] == ';'):
            condition_parts.append(current_part)
            current_part = ''
        else:
            current_part += string[i]
        i += 1
    
    # TODO klopt counting mechanisme
    condition_parts.append(current_part)
    return condition_parts

def handle_condition(condition):
    # TODO handle or / and / ...
    return 'A1>1'

def handle_if_logic(cells, excel_if):
    # Remove 'IF(' and last bracket ')'
    if_statement = excel_if[3:-1]
    
    parts = split_up_if_formula(if_statement)
    
    # Handle the condition statement
    condition = handle_condition(parts[0])
        
    excel_if = 'IF(' + condition + '){'
    for part in parts[1:]:
        cells, formula = ExcelHandler.extract_formula_cells(part, cells)
        excel_if += formula + '}else{'
    return cells, excel_if[-5]