from ExcelHandler.excel_extractor import extract_formula_cells
from ExcelHandler.excel_helpers import extract_col_row_from_excel_cell, is_sum
from Util.Cell import Cell



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
    arguments_unformatted = excel_formula.split(';')
    for arg in arguments_unformatted:
        
        # TODO handle recursion / nested functions
        # Max / min / ...
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


def handle_sum_calculation(cells, excel_sum):
    excel_sum = excel_sum[4:]
    formula = ''
    sums = get_sums(excel_sum)
    for sum in sums:
        cells, formula = extract_formula_cells(sum, cells, formula)
    return cells, formula