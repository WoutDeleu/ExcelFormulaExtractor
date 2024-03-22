import pandas as pd
import warnings
import formulas as fs
from util import *
from Cell import *
from excel_helpers import *
from excel_translator import *

warnings.simplefilter(action='ignore', category=UserWarning)


def resolve_cell(workbook, cell, formulas, values):
    sheet = workbook[cell.sheetname]
    print('Resolving cell: ' + cell.location + str(sheet[cell.location].value))
    if(isinstance(sheet[cell.location].value, int) or isinstance(sheet[cell.location].value, float)):
        print('Value')
        print('Cell: ' + cell.location + str(sheet[cell.location].value))
        values.append(CellValue(cell, sheet[cell.location].value))
        print()
        
        return formulas, values
    else:
        print('Formula')
        
        # library imported functions
        function = fs.Parser().ast(sheet[cell.location].value)[1].compile()
        print('Cells used using the shitty library: ' + str(list(function.inputs)))
        
        cells, formula = extract_formula_cells(sheet[cell.location].value)
        print('Cells used using the my own beautifull code: ' + list_to_string(cells))
        print('Translated formula: ' + formula)
        print()
        formulas.append(CellFormula(cell, formula))
        
        for cell in cells:
            formulas, values = resolve_cell(workbook, cell, formulas, values)
            
        return formulas, values
    
    
    
def main():
    workbook = read_in_excel('Draft PB-berekening - WERKVERSIE V4.xlsx')
    
    starting_cell = Cell('Tax Calculation', 'C41')
    starting_cell = Cell('Tax Calculation', 'C56')
    
    # Stack to keep track of formulas and values 
    formulas = []
    values = []
    
    formulas, values = resolve_cell(workbook, starting_cell, formulas, values)
    
    # for value in values:
    #     print(value.cell.location)
    #     print(value.value)
    # for value in formulas:
    #     print(value.cell.location)
    #     print(value.formula)
    
if __name__ == '__main__': 
    main()
