from ExcelHandler.ExcelFunctions.handle_date import handle_date, handle_datedif
from ExcelHandler.excel_helpers import *
from ExcelHandler.ExcelFunctions.handle_if import handle_if_logic
from ExcelHandler.ExcelFunctions.handle_round import handle_round
from ExcelHandler.ExcelFunctions.handle_sum_max_min import handle_sum_min_max
from ExcelHandler.ExcelFunctions.handle_vlookup import handle_vlookup
from ExcelHandler.handle_left_mid_right import handle_left, handle_mid, handle_right
from Util.util import is_letter_or_number
from Util.DataStructures import Queue, Set

def add_to_resulting_formula(resulting_formula, formula, operators):
    resulting_formula += formula
    if not operators.is_empty():
        resulting_formula += operators.pop()
    return resulting_formula


def extract_formula_cells(sheetname, excel_formula, formula='', cells=Set()):
    if excel_formula == '':
        return cells, formula
    
    # remove first character (=)
    if excel_formula[0] == '=':
        excel_formula = remove_char_from_string(excel_formula, 0)
    operators, parts = split_up_excel_formula(excel_formula)
    
    for element in parts.get_list():
        
        if is_iferror(element[:7]):
            cells, current_formula = extract_formula_cells(sheetname, element[8:-3], formula='', cells=cells)
            current_formula = '(iferror)' + current_formula
            
        elif is_if(element[:2]):
            cells, current_formula = handle_if_logic(cells, sheetname, element)
            
        elif is_sum(element[:3]) or is_max(element[:3]) or is_min(element[:3]):
            cells, current_formula = handle_sum_min_max(cells, sheetname, element)
            
        elif is_VLOOKUP(element[:7]):
            cells, current_formula = handle_vlookup(cells, sheetname, element)
        
        elif is_round(element[:5]):
            cells, current_formula = handle_round(cells, sheetname, element)
        
        elif is_datedif(element[:7]):
            cells, current_formula = handle_datedif(cells, sheetname, element)
        
        elif is_date(element[:4]):
            cells, current_formula = handle_date(cells, sheetname, element)
        
        elif is_left(element[:4]):
            cells, current_formula = handle_left(cells, sheetname, element)
            
        elif is_mid(element[:3]):
            cells, current_formula = handle_mid(cells, sheetname, element)
            
        elif is_right(element[:5]):
            cells, current_formula = handle_right(cells, sheetname, element)
            
        
        elif is_number(element) or is_percentage(element):
            current_formula = element
        
        elif is_excel_cell(element):
            cells.append(Cell(sheetname, element))
            current_formula = format_namespace(sheetname) + '_' + element
            
        elif is_fully_covered_by_brackets(element):
            cells, current_formula = extract_formula_cells(sheetname, element[1:-1], formula='', cells=cells)
            current_formula = '(' + current_formula + ')'
            
        elif element[0] == '+':
            cells, current_formula = extract_formula_cells(sheetname, element[1:], formula='', cells=cells)
            current_formula = '(' + current_formula + ')'
            
        elif element[0] == '-':
            cells, current_formula = extract_formula_cells(sheetname, element[1:], formula='', cells=cells)
            current_formula = '(' + element[0] + current_formula + ')'
            
        # TODO more extensive check?
        # Reference to another sheet
        elif is_led_by_sheetname(element):
            element = handle_leading_space(element)
            sheet_location_array = element.split('!')
            if sheet_location_array[1].find(" ") != -1:
                sheet_location_array[1] = sheet_location_array[1].replace(" ", "")
            relative_cell = handle_absolute_relative_cell(sheet_location_array[1])
            cells.append(Cell(sheet_location_array[0][1:-1], relative_cell))
            current_formula = format_namespace(sheet_location_array[0][1:-1]) + '_' + relative_cell
            
        elif element[0] == '\"':
            current_formula = element
        
        elif element[0] == ' ' or element == '' or element == ',' or element == ';' or element == ':' or element == '=' or element == None:  
            print('Invalid formula: ' + element)
            raise Exception('Invalid formula')
        
        # If is text/string value
        else:
            current_formula = '\"' + element + '\"'
        
        formula = add_to_resulting_formula(formula, current_formula, operators)
    
    return cells, formula


def split_up_excel_formula(string):
    # Queues
    operators = Queue()
    parts = Queue()
    
    current_part = ''
    brackets_to_close = 0
    quotes_to_close = 0
    
    is_allowed_to_close = False
    
    i = 0
    while i < len(string):
        if brackets_to_close > 0:
            current_part += string[i]
            if string[i] == ')':
                brackets_to_close -= 1
            if string[i] == '(':
                brackets_to_close += 1
                
        elif quotes_to_close > 0:
            current_part += string[i]
            if string[i] == '\'':
                quotes_to_close -= 1
                
        elif string[i] == '(':
            current_part += '('
            brackets_to_close += 1
            
        elif string[i] == ')':
            current_part += ')'
            brackets_to_close -= 1
            
        elif string[i] == '\'':
            current_part += '\''
            quotes_to_close += 1
            
        elif is_sum(string[i:i+3]):
            current_part += 'SUM('
            brackets_to_close += 1
            i += 3
            
        elif is_max(string[i:i+3]):
            current_part += 'MAX('
            brackets_to_close += 1
            i += 3
        
        elif is_min(string[i:i+3]):
            current_part += 'MIN('
            brackets_to_close += 1
            i += 3
            
        elif is_iferror(string[i:i+7]):
            current_part += 'IFERROR('
            brackets_to_close += 1
            i += 7
            
        elif is_if(string[i:i+2]):
            current_part += 'IF('
            brackets_to_close += 1
            i += 2
            
        elif is_round(string[i:i+5]):
            current_part += 'ROUND('
            brackets_to_close += 1
            i += 5
        
        elif is_VLOOKUP(string[i:i+7]):
            current_part += 'VLOOKUP('
            brackets_to_close += 1
            i += 7
            
        elif is_datedif(string[i:i+7]):
            current_part += 'DATEDIF('
            brackets_to_close += 1
            i += 7
        
        elif is_date(string[i:i+4]):
            current_part += 'DATE('
            brackets_to_close += 1
            i += 4
            
        elif is_left(string[i:i+4]):
            current_part += 'LEFT('
            brackets_to_close += 1
            i += 4
        
        elif is_mid(string[i:i+3]):
            current_part += 'MID('
            brackets_to_close += 1
            i += 3
            
        elif is_right(string[i:i+5]):
            current_part += 'RIGHT('
            brackets_to_close += 1
            i += 5
        
        elif is_operator(string[i]) and ((brackets_to_close == 0 and quotes_to_close == 0) or is_allowed_to_close) and current_part != '':
            operators.add(string[i])
            parts.add(current_part)
            current_part = ''
            is_allowed_to_close = False
            
        else:
            current_part += string[i]
            if is_letter_or_number(string[i]):
                is_allowed_to_close = True
        i += 1
    
    if brackets_to_close == 0:
        # TODO make more elegant - integrate in previous while loop
        parts.add(current_part)
    
    formatted_parts = Queue()
    for part in parts.get_list():
        part = handle_absolute_relative_cell(part)
        formatted_parts.add(part)
    
    return operators, formatted_parts

