import warnings
import sys
from Util.DataStructures import Stack
from Util.util import *
from Util.Cell import *
from ExcelHandler.excel_helpers import *
from ExcelHandler.excel_extractor import *
from tkinter.filedialog import askopenfilename

warnings.simplefilter(action='ignore', category=UserWarning)


def resolve_cell(workbook, cell, formulas, values):
    sheet = workbook[cell.sheetname]
    print('Resolving cell: ' + cell.location + " " + str(sheet[cell.location].value))

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
    filename = askopenfilename()
    workbook = read_in_excel(filename)
    starting_cell = Cell(sys.argv[1], sys.argv[2])
    
    # Stack to keep track of formulas and values 
    formulas = Stack()
    values = Stack()
    
    formulas, values = resolve_cell(workbook, starting_cell, formulas, values)
    
    print('######################################################################################################')
    print('##########################################  VALUES  ##################################################')
    print('######################################################################################################')
    for value in values.get_list():
        print(value.cell.location + ': ' + str(value.value))
    print()
    
    print('######################################################################################################')
    print('#########################################  FORMULAS  #################################################')
    print('######################################################################################################')
    for value in formulas.get_list():
        print(value.cell.location + ': ' + str(value.formula))
    
    
if __name__ == '__main__':
    main()
