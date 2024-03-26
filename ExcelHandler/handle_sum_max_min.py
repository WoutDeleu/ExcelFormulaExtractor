import ExcelHandler
from ExcelHandler.excel_helpers import extract_col_row_from_excel_cell, is_excel_range, is_max, is_min, is_sum
from Util.Cell import Cell



def split_up(excel_formula):
    parts = []
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
            parts.append(current_sum)
            current_sum = ''
        else:
            current_sum += ch
            
    parts.append(current_sum)
    return parts


def handle_range(sum_range, cells, formula, is_max_min):
    start_col, start_row = extract_col_row_from_excel_cell(sum_range.split(':')[0])
    end_col, end_row = extract_col_row_from_excel_cell(sum_range.split(':')[1])
    start_col_last_char = start_col[-1]
    start_col_except_last_char = start_col[:-1]
    end_col_last_char = end_col[-1]
    
    for i in range(int(start_row), int(end_row)+1):
        for j in range(ord(start_col_last_char), ord(end_col_last_char) + 1):
            cells.append(Cell('Tax Calculation', start_col_except_last_char + chr(j) + str(i)))
            if is_max_min:
                formula += start_col_except_last_char + chr(j) + str(i) + ';'
            else:
                formula += start_col_except_last_char + chr(j) + str(i) + '+'

    return cells, formula


def handle_sum_min_max(cells, excel):
    is_max_min = False
    if is_max(excel[:3]) or is_min(excel[:3]):
        is_max_min = True
        max_or_min = excel[:3]
        formula = max_or_min+'('
    else:
        formula = '('
    excel = excel[4:-1]
    parts = split_up(excel)
    for part in parts:
        if is_excel_range(part):
            cells, formula = handle_range(part, cells, formula, is_max_min)
        else:
            cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(part, formula, cells)
            if is_max_min:
                formula += ';'
            else:
                formula += '+'
    return cells, formula[:-1]+')'