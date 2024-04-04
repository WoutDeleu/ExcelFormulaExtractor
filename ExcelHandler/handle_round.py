import ExcelHandler
from ExcelHandler.excel_helpers import split_up_formulas


def handle_round(cells, sheetname, element):
    parts = split_up_formulas(element[6:-1])
    formula = 'round('
    cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[0],  formula=formula, cells=cells)
    cells, temp_formula = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[1], cells=cells)
    formula += ';' + temp_formula + ')'
    return cells, formula