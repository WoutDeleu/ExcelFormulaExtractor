import warnings
import sys
from Util.DataStructures import Stack
from Util.util import *
from Util.Cell import *
from ExcelHandler.excel_helpers import *
from ExcelHandler.excel_extractor import *
from tkinter.filedialog import askopenfilename
import datetime
from openpyxl.worksheet.formula import ArrayFormula

warnings.simplefilter(action='ignore', category=UserWarning)

def handle_constants(formulas, values, exceptions, sheet, cell):
    if is_constant(sheet[cell.location].value):
        value = sheet[cell.location].value
    else: 
        if is_int(sheet[cell.location].value[1:]):
            value = int(sheet[cell.location].value[1:])
        elif is_float(sheet[cell.location].value[1:]):
            value = float(sheet[cell.location].value[1:])
        
    print('Cell: ' + cell.location)
    print('Value: ' + str(sheet[cell.location].value))
    print()
    
    values.add(CellValue(cell, value))
    
    return formulas, values, exceptions


def handle_exceptions(formulas, values, exceptions, sheet, cell):
    if sheet[cell.location].value == "/":
        exceptions.add(CellValue(cell, '/'))
        return formulas, values, exceptions

    # TODO handle array functions
    if type(sheet[cell.location].value) == ArrayFormula:
        pass
    
    # TODO date time objects
    if isinstance(sheet[cell.location].value, datetime.time):
        pass
    
    if sheet[cell.location].value == None or sheet[cell.location].value == '':
        print('Cell: ' + cell.location)
        print('None')
        print()
        
        exceptions.add(CellValue(cell, 'None'))
        
    if sheet[cell.location].value == 'Ottignies- Louvain- La-Neuve':
        pass
    return formulas, values, exceptions
    

def handle_formulas(formulas, values, exceptions, workbook, sheet, cell):
    cells, formula = extract_formula_cells(cell.sheetname, sheet[cell.location].value, cells=Set())
    
    
    print('Cells: ' + list_to_string(cells.get_list()))
    print('Translated formula: ' + str(formula))
    print()
    
    formulas.add(CellFormula(cell, formula))
    
    for cell in cells.get_list():
        if not formulas.contains(cell) and not values.contains(cell) and not exceptions.contains(cell) and cell.location != '#REF':
            formulas, values, exceptions = resolve_cell(workbook, cell, formulas, values, exceptions)
        
        elif formulas.contains(cell):
            formulas.move_to_top(cell)
        
        elif values.contains(cell):
            values.move_to_top(cell)
        
        elif exceptions.contains(cell):
            exceptions.move_to_top(cell)
        
    return formulas, values, exceptions

def resolve_cell(workbook, cell, formulas, values, exceptions):    
    sheet = workbook[cell.sheetname]
    print('Resolving cell: ' + cell.location + " " + str(sheet[cell.location].value))

    if is_exception(sheet[cell.location].value):
        return handle_exceptions(formulas, values, exceptions, sheet, cell)
    
    elif is_constant(sheet[cell.location].value) or (sheet[cell.location].value[0] == '=' and is_int(str(sheet[cell.location].value[1:]) or is_float(sheet[cell.location].value)[1:])):
        return handle_constants(formulas, values, exceptions, sheet, cell)
    
    else:
        return handle_formulas(formulas, values, exceptions, workbook, sheet, cell)
    


def run_full_analysis(workbook):
    starting_cells = ['C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38', 'C40', 'C41', 'C42', 'C43', 'C44', 'C45', 'C46', 'C47', 'C48', 'C49', 'C50', 'C51', 'C52', 'C53', 'C54', 'C55', 'C56', 'C57', 'C58', 'C60', 'C61', 'D32', 'D33', 'D34', 'D35', 'D36', 'D37', 'D38', 'D40', 'D41', 'D42', 'D43', 'D44', 'D45', 'D46', 'D47', 'D48', 'D49', 'D50', 'D51', 'D52', 'D53', 'D54', 'D55', 'D56', 'D57', 'D58', 'D60']
    for starting_cell in starting_cells:
        starting_cell = Cell('Tax Calculation', starting_cell)
        
        # Stack to keep track of formulas and values 
        formulas = Stack()
        values = Stack()
        exceptions = Stack()
        
        formulas, values, exceptions = resolve_cell(workbook, starting_cell, formulas, values, exceptions)
        
        print_results(formulas, values, exceptions)


def main():
    filename = askopenfilename()
    workbook = read_in_excel(filename)
    
    # TODO - Remove this hardcoded list of starting cells
    run_full_analysis(workbook)
    
    # starting_cell = Cell(sys.argv[1], sys.argv[2])
    
    # # Stack to keep track of formulas and values 
    # formulas = Stack()
    # values = Stack()
    # exceptions = Stack()
    
    # formulas, values, exceptions = resolve_cell(workbook, starting_cell, formulas, values, exceptions)
    # print_results(formulas, values, exceptions)
    
    
if __name__ == '__main__':
    main()



