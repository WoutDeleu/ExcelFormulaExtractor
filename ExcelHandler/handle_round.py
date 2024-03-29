from ExcelHandler.excel_extractor import extract_formula_cells
from ExcelHandler.excel_helpers import split_up_formulas


def handle_round(cells, sheetname, element):
    parts = split_up_formulas(element[6:-1])
    cells, formula = extract_formula_cells(sheetname, parts[0], cells=cells)
    formula += '.round('
    cells, round = extract_formula_cells(sheetname, parts[1], cells=cells)
    formula += round + ')'
    return cells, formula