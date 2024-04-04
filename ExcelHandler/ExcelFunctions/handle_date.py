import ExcelHandler
from ExcelHandler.excel_helpers import split_up_formulas


def handle_datedif(cells, sheetname, excel):
    parts = split_up_formulas(excel[8:-1])
    
    cells, date1 = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[0],  formula='', cells=cells)
    cells, date2 = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[1],  formula='', cells=cells)
    
    cells, part = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, parts[2],  formula='', cells=cells)
    part = part.lower().replace('"', '')
    
    # TODO handle datetime
    if part == 'y':
        pass
    elif part == 'm':
        pass
    elif part == 'd':
        pass
    elif part == 'md':
        pass
    elif part == 'ym':
        pass
    elif part == 'yd':
        pass
    else:
        print('Error in datedif')
        raise ValueError('Error in datedif')
    
    return cells, 'DATEDIF(' + date1 + ';' + date2 + ';' + part + ')'


def handle_date(cells, sheetname, excel):
    parts = split_up_formulas(excel[5:-1])
    resulting_formula = 'DATE('
    for part in parts:
        cells, formula = ExcelHandler.excel_extractor.extract_formula_cells(sheetname, part,  formula='', cells=cells)
        resulting_formula += formula + ';'
        
    return cells, resulting_formula[:-1] + ')'