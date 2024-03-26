import ExcelHandler
from ExcelHandler.excel_helpers import extract_col_row_from_excel_cell, is_excel_range, is_sum
from Util.Cell import Cell



def split_up_sum(excel_formula):
    sums = []
    current_sum = ''
    counter_closing_brackets_needed = 0
    for ch in excel_formula:
        if ch == '(':
            counter_closing_brackets_needed += 1
            current_sum += ch
        elif ch == ')':
            counter_closing_brackets_needed -= 1
            current_sum += ch
        elif ch == ';'  and counter_closing_brackets_needed == 0:
            sums.append(current_sum)
            current_sum = ''
        else:
            current_sum += ch
            
    sums.append(current_sum)
    return sums


def handle_range(sum_range, cells, formula):
    start_col, start_row = extract_col_row_from_excel_cell(sum_range.split(':')[0])
    end_col, end_row = extract_col_row_from_excel_cell(sum_range.split(':')[1])
    start_col_last_char = start_col[-1]
    start_col_except_last_char = start_col[:-1]
    end_col_last_char = end_col[-1]
    
    for i in range(int(start_row), int(end_row)+1):
        for j in range(ord(start_col_last_char), ord(end_col_last_char) + 1):
            cells.append(Cell('Tax Calculation', start_col_except_last_char + chr(j) + str(i)))
            formula += start_col_except_last_char + chr(j) + str(i) + '+'

    return cells, formula


def handle_sum(cells, excel_sum):
    excel_sum = excel_sum[4:-1]
    formula = '('
    sums = split_up_sum(excel_sum)
    for sum in sums:
        if is_excel_range(sum):
            cells, formula = handle_range(sum, cells, formula)
        else:
            cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(sum, formula, cells)
            formula += '+'
    return cells, formula[:-1]+')'