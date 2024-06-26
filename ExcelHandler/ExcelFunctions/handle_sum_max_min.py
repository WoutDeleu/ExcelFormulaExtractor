import ExcelHandler
from ExcelHandler.excel_helpers import extract_col_row_from_excel_cell, handle_absolute_relative_cell, is_absolute_reference, is_excel_cell, is_excel_range, is_max, is_min, is_sum, split_up_formulas
from Util.Cell import Cell
from Util.util import format_namespace


def handle_range(sum_range, cells, formula, is_max_min, sheetname):
    start_range = sum_range.split(':')[0]
    end_range = sum_range.split(':')[1]
    
    # TODO more extensive check?
    # TODO in theory sheets should be the same
    if start_range[0] == '\'':
        sheet_location_array = start_range.split('!')
        sheetname = sheet_location_array[0][1:-1]
        start_range = sheet_location_array[1]
    if end_range[0] == '\'':
        sheet_location_array = end_range.split('!')
        sheetname = sheet_location_array[0][1:-1]
        end_range = sheet_location_array[1]
    
    start_range = handle_absolute_relative_cell(start_range)
    end_range = handle_absolute_relative_cell(end_range)
    
    start_col, start_row = extract_col_row_from_excel_cell(start_range)
    end_col, end_row = extract_col_row_from_excel_cell(end_range)
    start_col_last_char = start_col[-1]
    start_col_except_last_char = start_col[:-1]
    end_col_last_char = end_col[-1]
    
    for i in range(int(start_row), int(end_row)+1):
        for j in range(ord(start_col_last_char), ord(end_col_last_char) + 1):
            cells.append(Cell(sheetname, start_col_except_last_char + chr(j) + str(i)))
            if is_max_min:
                formula += format_namespace(sheetname) + '_' + start_col_except_last_char + chr(j) + str(i) + ';'
            else:
                formula += format_namespace(sheetname) + '_' + start_col_except_last_char + chr(j) + str(i) + '+'

    return cells, formula


def handle_sum_min_max(cells, sheetname, excel):
    is_max_min = False
    if is_max(excel[:3]) or is_min(excel[:3]):
        is_max_min = True
        max_or_min = excel[:3]
        formula = max_or_min+'('
    else:
        formula = '('
    excel = excel[4:-1]
    parts = split_up_formulas(excel)
    for part in parts:
        if is_excel_range(part):
            cells, formula = handle_range(part, cells, formula, is_max_min, sheetname)
        else:
            cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part, formula, cells)
            if is_max_min:
                formula += ';'
            else:
                formula += '+'
    return cells, formula[:-1]+')'