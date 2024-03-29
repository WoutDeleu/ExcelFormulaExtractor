import ExcelHandler
from ExcelHandler.excel_helpers import is_and, is_or, is_xor, is_not, split_up_conditions, split_up_formulas
from Util.Cell import Cell


def handle_condition(sheetname, condition, cells):
    if is_or(condition[:2]):
        parts = split_up_formulas(condition[3:-1])
        current_formula = ''
        for part in parts:
            cells, formula = handle_condition(sheetname, part, cells)
            current_formula += formula + '|'
        current_formula = '(' + current_formula[:-1] + ')'
        
    elif is_and(condition[:3]):
        parts = split_up_formulas(condition[4:-1])
        current_formula = ''
        for part in parts:
            cells, formula = handle_condition(sheetname, part, cells)
            current_formula += formula + '&'
        current_formula =  '(' + current_formula[:-1] + ')'
        
    elif is_xor(condition[:3]):
        parts = split_up_formulas(condition[4:-1])
        current_formula = ''
        for part in parts:
            cells, formula = handle_condition(sheetname, part, cells)
            current_formula += formula + '^'
        current_formula =  '(' + current_formula[:-1] + ')'
    elif is_not(condition[:3]):
        formula = condition[4:-1]
        cells, formula = handle_condition(sheetname, formula, cells)
        current_formula = '!(' + formula + ')'
    elif condition[0] == '(':
        cells, current_formula = handle_condition(sheetname, condition[1:-1], cells)
    else:
        parts, operators = split_up_conditions(condition)
        current_formula = ''
        for part in parts:
            cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part, cells=cells)
            current_formula += formula
            if operators:
                current_formula += operators.pop()
    return cells, current_formula
    


def handle_if_logic(cells, sheetname, excel_if):
    # Remove 'IF(' and last bracket ')'
    if_statement = excel_if[3:-1]
    
    parts = split_up_formulas(if_statement)
    
    # Handle the condition statement
    cells, condition = handle_condition(sheetname, parts[0], cells)
        
    excel_if = 'if(' + condition + '){'
    for part in parts[1:]:
        cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part, cells=cells)
        excel_if += formula + '}else{'
    return cells, excel_if[:-5]