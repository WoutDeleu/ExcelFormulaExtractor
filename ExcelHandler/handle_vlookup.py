import ExcelHandler
from ExcelHandler.excel_helpers import split_up_formulas
from ExcelHandler.handle_sum_max_min import handle_range

def handle_vlookup(cells, sheetname, element):
    parts = split_up_formulas(element[8:-1])
    
    cells, matching_value = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[0],  cells=cells)
    
    cells, range = handle_range(parts[1], cells, '', True, sheetname)
    
    offset = parts[2]
    
    formula = 'for(cell in ' + '[' + range + ']{'+ 'if(cell == '+ matching_value + ')' + 'return update_column(cell, ' + offset + ')' +  '}' + '}'

    return cells, formula