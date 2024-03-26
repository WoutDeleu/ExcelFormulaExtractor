from Util.Cell import Cell

def split_up_if_formula(string):
    condition_parts = []
    
    condition_is_completed = False
    brackets_to_close = 0
    
    i = 0
    current_part = ''
    while not condition_is_completed:
        if string[i] == '(':
            brackets_to_close += 1
        if string[i] == ')':
            brackets_to_close -= 1
        
        if brackets_to_close == 0 and (string[i] == ',' or string[i] == ';'):
            condition_parts.append(current_part)
            current_part = ''
            condition_is_completed = True
        else:
            current_part += string[i]
        i += 1
    
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
    return condition_parts

def handle_if_logic(cells, excel_if):
    # Remove 'IF(' and last bracket ')'
    if_statement = excel_if[3:-1]
    parts = split_up_if_formula(if_statement)
    
    # Handle if
    # for part in parts:
        # TODO get cells
        # TODO get formula
        
    # TODO handle or / and / ...
    
    # return cells, excel_if