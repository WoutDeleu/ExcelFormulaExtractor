from ExcelHandler.excel_helpers import *
from ExcelHandler.handle_if import handle_if_logic
from ExcelHandler.handle_round import handle_round
from ExcelHandler.handle_sum_max_min import handle_sum_min_max
from ExcelHandler.handle_vlookup import handle_vlookup
from Util.util import is_letter_or_number
from Util.DataStructures import Queue, Set
from openpyxl.worksheet.formula import ArrayFormula

def add_to_resulting_formula(resulting_formula, formula, operators):
    resulting_formula += formula
    if not operators.is_empty():
        resulting_formula += operators.pop()
    return resulting_formula


def extract_formula_cells(sheetname, excel_formula, formula='', cells=Set()):
    # TODO handle references to other sheets - nu, default sheet is 'Tax Calculation'
    # remove first character (=)
    if excel_formula == None or excel_formula == '':
        return cells, ''
    
    # TODO handle array functions
    if type(excel_formula) == ArrayFormula:
        return cells, ''
    
    if excel_formula[0] == '=':
        excel_formula = remove_char_from_string(excel_formula, 0)

    operators, parts = split_up_excel_formula(excel_formula)
    
    for element in parts.get_list():
        if is_iferror(element[:7]):      
            element = element[7:-1]
            element = split_up_formulas(element)[0]
            
        if is_sum(element[:3]) or is_max(element[:3]) or is_min(element[:3]):
            cells, current_formula = handle_sum_min_max(cells, sheetname, element)
            
        elif is_if(element[:2]):
            cells, current_formula = handle_if_logic(cells, sheetname, element)
            
        elif is_VLOOKUP(element[:7]):
            cells, current_formula = handle_vlookup(cells, sheetname, element)
        
        elif is_round(element[:5]):
            cells, current_formula = handle_round(cells, sheetname, element)
        
        elif is_number(element):
            # TODO extra - extra aanduiding voor een getal
            current_formula = element
        
        elif is_percentage(element):
            # TODO extra - extra aanduiding voor een getal
            current_formula = element
        
        elif is_excel_cell(element):
            cells.append(Cell(sheetname, element))
            current_formula = format_namespace(sheetname) + '_' + element
        
        elif element[0] == '-':
            cells, current_formula = extract_formula_cells(sheetname, element[1:], formula=formula, cells=cells)
            current_formula = '(-' + current_formula + ')'
            
        elif element[0] == '(':
            cells, current_formula = extract_formula_cells(sheetname, element[1:-1], formula=formula, cells=cells)
            current_formula = '(' + current_formula + ')'
        
        # Reference to another sheet
        # TODO more extensive check?
        elif element[0] == '\'':
            sheet_location_array = element.split('!')
            cells.append(Cell(sheet_location_array[0][1:-1], sheet_location_array[1]))
            current_formula = format_namespace(sheetname) + '_' + sheet_location_array[1]
        
        elif element[0] == '\"':
            current_formula = element
        
        elif element[0] == ' ' or element == '' or element == ',' or element == ';' or element == ':' or element == '=' or element == None:  
            # print('Invalid formula: ' + element)
            # raise Exception('Invalid formula')
            formula = ''
        
        # If is text/string value
        else:
            current_formula = '\"' +element + '\"'
        
        formula = add_to_resulting_formula(formula, current_formula, operators)
    
    return cells, formula



def split_up_excel_formula(string):
    # TODO fout!!!!!!
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
                
        elif string[i] == '(':
            current_part += '('
            brackets_to_close += 1
            
        elif string[i] == ')':
            current_part += ')'
            brackets_to_close -= 1
            
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
            
        elif is_if(string[i:i+2]):
            current_part += 'IF('
            brackets_to_close += 1
            i += 2
            
        elif is_iferror(string[i:i+7]):
            current_part += 'IFERROR('
            brackets_to_close += 1
            i += 7
            
        elif is_round(string[i:i+5]):
            current_part += 'ROUND('
            brackets_to_close += 1
            i += 5
        
        elif is_VLOOKUP(string[i:i+7]):
            current_part += 'VLOOKUP('
            brackets_to_close += 1
            i += 7
        
        elif is_operator(string[i]) and (brackets_to_close == 0 or is_allowed_to_close) and current_part != '':
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
        
    return operators, parts

