import ExcelHandler
from ExcelHandler.excel_helpers import is_excel_range, split_up_formulas
from ExcelHandler.handle_sum_max_min import handle_range

def handle_vlookup(cells, sheetname, element):
    """Handles the VLOOKUP formula.

    Args:
        cells (dict): A dictionary containing the cell values.
        sheetname (str): The name of the sheet.
        element (str): The VLOOKUP formula.

    Returns:
        tuple: A tuple containing the updated cells dictionary and the formula.

    Warning:
        Approximate match is currently not supported. It chouses the exact match.
    """
    parts = split_up_formulas(element[8:-1])
    
    cells, matching_value = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[0],  cells=cells)
    
    # TODO what if not a range???
    if is_excel_range(parts[1]):
        cells, range = handle_range(parts[1], cells, '', True, sheetname)
        range = range[:-1]
    else :
        cells, range = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[1], cells=cells)
    
    offset = parts[2]
    
    formula = 'for(cell in ' + '[' + range + ']{'+ 'if(cell == '+ matching_value + '){' + 'return update_column(cell, ' + offset + ')' +  '}' + '}'

    return cells, formula