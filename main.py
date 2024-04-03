import warnings
import sys
from Util.DataStructures import Stack
from Util.util import *
from Util.Cell import *
from ExcelHandler.excel_helpers import *
from ExcelHandler.excel_extractor import *
from tkinter.filedialog import askopenfilename

warnings.simplefilter(action='ignore', category=UserWarning)


def resolve_cell(workbook, cell, formulas, values, errors):
    sheet = workbook[cell.sheetname]
    print('Resolving cell: ' + cell.location + " " + str(sheet[cell.location].value))

    if sheet[cell.location].value == None:
        print('Cell: ' + cell.location)
        print('None')
        
        errors.add(CellValue(cell, 'None'))
        print()
        
        return formulas, values, errors
        
    elif is_constant(sheet[cell.location].value):
        print('Cell: ' + cell.location)
        print('Value: ' + str(sheet[cell.location].value))
        
        values.add(CellValue(cell, sheet[cell.location].value))
        print()
        
        return formulas, values, errors
    
    elif sheet[cell.location].value[0] == '=' and (is_int(sheet[cell.location].value[1:]) or is_float(sheet[cell.location].value[1:])):
        if is_int(sheet[cell.location].value[1:]):
            value = int(sheet[cell.location].value[1:])
        elif is_float(sheet[cell.location].value[1:]):
            value = float(sheet[cell.location].value[1:])
            
        print('Cell: ' + cell.location)
        print('Value: ' + str(sheet[cell.location].value))
        
        values.add(CellValue(cell, value))
        print()
        
        return formulas, values, errors
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
            formulas, values, errors = resolve_cell(workbook, cell, formulas, values, errors)
            
        return formulas, values, errors
    
    
def main():
    filename = askopenfilename()
    workbook = read_in_excel(filename)
    starting_cell = Cell(sys.argv[1], sys.argv[2])
    
    # Stack to keep track of formulas and values 
    formulas = Stack()
    values = Stack()
    errors = Stack()
    
    formulas, values, errors = resolve_cell(workbook, starting_cell, formulas, values, errors)
    
    print('######################################################################################################')
    print('##########################################  VALUES  ##################################################')
    print('######################################################################################################')
    for value in values.get_list():
        print(value.cell.location + ': ' + str(value.value))
        # print(value.cell.location + '-' + value.cell.sheetname + ': ' + str(value.value))
    print()
    
    print('######################################################################################################')
    print('#########################################  FORMULAS  #################################################')
    print('######################################################################################################')
    for value in formulas.get_list():
        print(value.cell.location + ': ' + str(value.formula))
        
    print('######################################################################################################')
    print('##########################################  ERRORS  ##################################################')
    print('######################################################################################################')
    for error in errors.get_list():
        print(error.cell.location + ': ' + str(error.value))
    
    
if __name__ == '__main__':
    main()
