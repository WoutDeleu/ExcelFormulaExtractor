import unittest
import pandas as pd
import warnings
import formulas as fs
from Util.DataStructures import Stack
from Util.util import *
from Util.Cell import *
from ExcelHandler.excel_helpers import *
from ExcelHandler.excel_extractor import *

warnings.simplefilter(action='ignore', category=UserWarning)


def resolve_cell(workbook, cell, formulas, values):
    sheet = workbook[cell.sheetname]
    print('Resolving cell: ' + cell.location + " " + str(sheet[cell.location].value))
    
    if '=IF(L520=0,0,-L520*(M515/(M515+Q515)))' in str(sheet[cell.location].value):
        print()
    
    if(isinstance(sheet[cell.location].value, int) or isinstance(sheet[cell.location].value, float)):
        print('Cell: ' + cell.location + ' ' +  str(sheet[cell.location].value))
        
        values.add(CellValue(cell, sheet[cell.location].value))
        print()
        
        return formulas, values
    else:
        if sheet[cell.location].value != None:
            cells, formula = extract_formula_cells(cell.sheetname, sheet[cell.location].value, cells=Set())
        
        else:
            formula = 0
            cells = Set()
        
        print('Cells: ' + list_to_string(cells.get_list()))
        print('Translated formula: ' + str(formula))
        print()
            
        formulas.add(CellFormula(cell, formula))
        
        for cell in cells.get_list():
            formulas, values = resolve_cell(workbook, cell, formulas, values)
            
        return formulas, values
    
    
    
def main():

    workbook = read_in_excel('Draft PB-berekening - WERKVERSIE V4.xlsx')
    
    starting_cell = Cell('Tax Calculation', 'C41')
    starting_cell = Cell('Tax Calculation', 'C56')
    
    # Stack to keep track of formulas and values 
    formulas = Stack()
    values = Stack()
    
    formulas, values = resolve_cell(workbook, starting_cell, formulas, values)
    
    
    for value in values:
        print(value.cell.location)
        print(value.value)
    for value in formulas:
        print(value.cell.location)
        print(value.formula)
    
if __name__ == '__main__':
    main()
