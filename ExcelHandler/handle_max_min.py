from ExcelHandler.excel_helpers import extract_col_row_from_excel_cell
from Util.Cell import Cell


def get_max_mins_elements(excel_formula):
    max_mins = []
    counter_closing_brackets_needed = 1
    current_max_min = ''
    for ch in excel_formula:
        if ch == '(':
            counter_closing_brackets_needed += 1
        elif ch == ')':
            counter_closing_brackets_needed -= 1
            if counter_closing_brackets_needed == 0:
                max_mins.append(current_max_min)
                current_max_min = ''
        else:
            current_max_min += ch
    return max_mins


def handle_max_min_logic(excel_formula, cells, formula, max_or_min):
    arguments_unformatted = excel_formula.split(';')
    for arg in arguments_unformatted:
        
        # TODO handle recursion / nested functions
        # Max / min / ...
        if ':' in arg:
            start_col, start_row = extract_col_row_from_excel_cell(arg.split(':')[0])
            end_col, end_row = extract_col_row_from_excel_cell(arg.split(':')[1])
            start_col_last_char = start_col[-1]
            start_col_except_last_char = start_col[:-1]
            end_col_last_char = end_col[-1]
            
            for i in range(int(start_row), int(end_row)+1):
                for j in range(ord(start_col_last_char), ord(end_col_last_char) + 1):
                    cells.append(Cell('Tax Calculation', start_col_except_last_char + chr(j) + str(i)))
                    formula += start_col_except_last_char + chr(j) + str(i) + ';'

        else:
            cells.append(Cell('Tax Calculation', arg))
            formula += arg + ';'
    
    return cells, max_or_min+formula[:-1]+')'


def handle_max_min(cells, excel_max_min):
    max_or_min = excel_max_min[:4]
    excel_max_min = excel_max_min[4:]
    formula = ''
    max_mins = get_max_mins_elements(excel_max_min)
    
    for arg in max_mins:
        cells, formula = handle_max_min_logic(arg, cells, formula, max_or_min)
    
    return cells, formula