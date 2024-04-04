import ExcelHandler
from ExcelHandler.excel_helpers import split_up_formulas


def handle_left(cells, sheetname, element):
    parts = split_up_formulas(element[5:-1])
    formula = 'LEFT('
    for part in parts:
        cells, f = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part,  formula='', cells=cells)
        formula += f + ';'
    return cells, formula[:-1] + ')'

def handle_mid(cells, sheetname, element):
    parts = split_up_formulas(element[5:-1])
    formula = 'MID('
    for part in parts:
        cells, f = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part,  formula='', cells=cells)
        formula += f + ';'
    return cells, formula[:-1] + ')'

def handle_right(cells, sheetname, element):
    parts = split_up_formulas(element[5:-1])
    formula = 'RIGHT('
    for part in parts:
        cells, f = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part,  formula='', cells=cells)
        formula += f + ';'
    return cells, formula[:-1] + ')'