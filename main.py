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
    if cell.location == 'C34':
        pass

    if sheet[cell.location].value == None:
        print('Cell: ' + cell.location)
        print('None')
        print()
        
        errors.add(CellValue(cell, 'None'))
        
        return formulas, values, errors
        
    elif is_constant(sheet[cell.location].value):
        print('Cell: ' + cell.location)
        print('Value: ' + str(sheet[cell.location].value))
        print()
        
        values.add(CellValue(cell, sheet[cell.location].value))
        
        return formulas, values, errors
    
    elif sheet[cell.location].value[0] == '=' and (is_int(sheet[cell.location].value[1:]) or is_float(sheet[cell.location].value[1:])):
        if is_int(sheet[cell.location].value[1:]):
            value = int(sheet[cell.location].value[1:])
        elif is_float(sheet[cell.location].value[1:]):
            value = float(sheet[cell.location].value[1:])
            
        print('Cell: ' + cell.location)
        print('Value: ' + str(sheet[cell.location].value))
        print()
        
        values.add(CellValue(cell, value))
        
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
            if not formulas.contains(cell) and not values.contains(cell) and not errors.contains(cell):
                formulas, values, errors = resolve_cell(workbook, cell, formulas, values, errors)
            
            elif formulas.contains(cell):
                formulas.move_to_top(cell)
            
            elif values.contains(cell):
                values.move_to_top(cell)
            
            elif errors.contains(cell):
                errors.move_to_top(cell)
            
        return formulas, values, errors
    
    
def run_full_analysis(workbook):
    starting_cells = ['C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38', 'C40', 'C41', 'C42', 'C43', 'C44', 'C45', 'C46', 'C47', 'C48', 'C49', 'C50', 'C51', 'C52', 'C53', 'C54', 'C55', 'C56', 'C57', 'C58', 'C60', 'C61', 'D32', 'D33', 'D34', 'D35', 'D36', 'D37', 'D38', 'D40', 'D41', 'D42', 'D43', 'D44', 'D45', 'D46', 'D47', 'D48', 'D49', 'D50', 'D51', 'D52', 'D53', 'D54', 'D55', 'D56', 'D57', 'D58', 'D60']
    for starting_cell in starting_cells:
        starting_cell = Cell('Tax Calculation', starting_cell)
        
        # Stack to keep track of formulas and values 
        formulas = Stack()
        values = Stack()
        errors = Stack()
        
        formulas, values, errors = resolve_cell(workbook, starting_cell, formulas, values, errors)
        
        print_results(formulas, values, errors)
        
        
def main():
    filename = askopenfilename()
    workbook = read_in_excel(filename)
    
    # TODO - Remove this hardcoded list of starting cells
    run_full_analysis(workbook)
    
    starting_cell = Cell(sys.argv[1], sys.argv[2])
    
    # Stack to keep track of formulas and values 
    formulas = Stack()
    values = Stack()
    errors = Stack()
    
    # formulas, values, errors = resolve_cell(workbook, starting_cell, formulas, values, errors)
    print_results(formulas, values, errors)
    
    
if __name__ == '__main__':
    main()
