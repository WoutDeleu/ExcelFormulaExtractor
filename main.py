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
import argparse

warnings.simplefilter(action='ignore', category=UserWarning)

# TODO - Remove this hardcoded list of starting cells
global table_cells
table_cells = ['C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38', 'C40', 'C41', 'C42', 'C43', 'C44', 'C45', 'C46', 'C47', 'C48', 'C49', 'C50', 'C51', 'C52', 'C53', 'C54', 'C55', 'C56', 'C57', 'C58', 'C60', 'C61', 'C63', 'D32', 'D33', 'D34', 'D35', 'D36', 'D37', 'D38', 'D40', 'D41', 'D42', 'D43', 'D44', 'D45', 'D46', 'D47', 'D48', 'D49', 'D50', 'D51', 'D52', 'D53', 'D54', 'D55', 'D56', 'D57', 'D58', 'D60', 'D113']
nested_table_cells = {'C32':'M524', 'D32': 'Q524', 'C34':'M166', 'D34':'Q166', 'C35': 'M581', 'D35': 'Q581', 'C36':'M526', 'D36': 'Q526', 'C37': 'M199', 'D37': 'Q199', 'C42': 'F601', 'D42': 'M601', 'C43': 'F643', 'D43': 'M643', 'C47': 'M1183', 'D47': 'Q1183', 'C48': 'M1188', 'D48': 'Q1188', 'C51': 'M1254', 'D51': 'Q1254', 'C58': '1441', 'D58': 'F1441' }


def is_already_calculated_externally(cell):
    if cell.sheetname != 'Tax Calculation':
        return False
    for starting_cell in table_cells:
        if cell.location == starting_cell:
            return True
    for key, value in nested_table_cells.items():
        if cell.location == value:
            return True
    return False

def handle_constants(formulas, values, exceptions, sheet, cell):
    if isinstance(sheet[cell.location].value, datetime.time) or isinstance(sheet[cell.location].value, datetime.datetime) or is_constant(sheet[cell.location].value):
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
        exceptions.add(CellValue(cell, sheet[cell.location].value))
    
    if sheet[cell.location].value == None or sheet[cell.location].value == '':
        print('Cell: ' + cell.location)
        print('None')
        print()
        
        exceptions.add(CellValue(cell, 'None'))
        
    if sheet[cell.location].value == 'Ottignies- Louvain- La-Neuve':
        exceptions.add(CellValue(cell, sheet[cell.location].value))
        print('Cell: ' + cell.location)
        print('Date: ' + str(sheet[cell.location].value))
        print()
        
    return formulas, values, exceptions
    

def handle_formulas(formulas, values, exceptions, workbook, sheet, cell):
    cells, formula = extract_formula_cells(cell.sheetname, sheet[cell.location].value, cells=Set())
    
    
    print('Cells: ' + list_to_string(cells.get_list()))
    print('Translated formula: ' + str(formula))
    print()
    
    formulas.add(CellFormula(cell, formula))
    
    for cell in cells.get_list():
        if cell.location == 'D113':
            pass
        # TODO - Remove this hardcoded list of starting cells
        if not formulas.contains(cell) and not values.contains(cell) and not exceptions.contains(cell) and cell.location != '#REF' and (not is_already_calculated_externally(cell)):
            formulas, values, exceptions = resolve_cell(workbook, cell, formulas, values, exceptions)
        
        elif cell.location == '#REF':
            exceptions.add(CellValue(cell, '#REF'))
            print('Cell: ' + cell.location)
            print('Value: #REF')
            print()
            
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
    
    elif is_constant(sheet[cell.location].value) or (sheet[cell.location].value[0] == '=' and is_int(str(sheet[cell.location].value[1:]) or is_float(sheet[cell.location].value)[1:])) or isinstance(sheet[cell.location].value, datetime.time) or isinstance(sheet[cell.location].value, datetime.datetime):
        return handle_constants(formulas, values, exceptions, sheet, cell)
    
    else:
        return handle_formulas(formulas, values, exceptions, workbook, sheet, cell)



def run_full_analysis(workbook):
    starting_cells = ['C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38', 'C40', 'C41', 'C42', 'C43', 'C44', 'C45', 'C46', 'C47', 'C48', 'C49', 'C50', 'C51', 'C52', 'C53', 'C54', 'C55', 'C56', 'C57', 'C58', 'C60', 'C61', 'C63', 'D32', 'D33', 'D34', 'D35', 'D36', 'D37', 'D38', 'D40', 'D41', 'D42', 'D43', 'D44', 'D45', 'D46', 'D47', 'D48', 'D49', 'D50', 'D51', 'D52', 'D53', 'D54', 'D55', 'D56', 'D57', 'D58', 'D60']
    names = ['prof_income', 'replacement_income', 'real_estate_income', 'miscellaneous_income', 'marriage_coefficient', 'movable_income', 'total_net_income', 'alimony_spec_soc_sec', 'joint_taxable_income', 'base_tax', 'min_tax_free_sum', 'applicable_tax_bef', 'min_foreign_exempt_income', 'principal_amount', 'state_tax', 'autonomy_factor', 'reduced_state_tax', 'federal_tax_deductions', 'plus_regional_surcharges', 'regional_tax_deductions', 'separate_taks', 'total_tax_1', 'witholding_taks', 'tax_credits', 'tax_increase_reversal', 'communal_tax', 'total_tax_2', 'spec_soc_sec_contr', 'balance_to_be_paid']
    i = 0
    for starting_cell in starting_cells:
        global nested_table_cells
        has_popped = False
        if starting_cell in nested_table_cells:
            popped_value = nested_table_cells.pop(starting_cell)
            has_popped = True
            
        
        global table_cells
        index = table_cells.index(starting_cell)
        table_cells.remove(starting_cell)
        
        starting_cell = Cell('Tax Calculation', starting_cell)
        
        # Stack to keep track of formulas and values 
        formulas = Stack()
        values = Stack()
        exceptions = Stack()
        
        formulas, values, exceptions = resolve_cell(workbook, starting_cell, formulas, values, exceptions)
        table_cells.insert(index, starting_cell.location)
        if has_popped:
            nested_table_cells.update({starting_cell.location: popped_value})
        print('Starting cell: ' + starting_cell.location)
        if starting_cell.location[0] == 'C':
            filename = names[i] + '_oldest'
        else:
            filename = names[i] + '_youngest'
        print_results(formulas, values, exceptions, filename=filename, write_to_file=True)
        i += 1
        if i == len(names):
            i = 0


def main(args):
    print('Welcome to the Excel Extraction Tool')
    if args.file is None:
        print('Please select the Excel file you want to analyse')
        filename = askopenfilename()
        workbook = read_in_excel(filename)
    else:
        workbook = read_in_excel(args.file)
    
    if args.full_analysis and args.single_cell:
        print('Please provide either the full analysis or the single cell analysis, not both at the same time')
        sys.exit()
    
    if args.full_analysis:
        run_full_analysis(workbook)
    else:
        if not args.single_cell:
            doese_run_full_analysis = input('Do you want the full analysis? (y/n):')
            if doese_run_full_analysis == 'y':
                run_full_analysis(workbook)

        if bool(args.cell is not None) != bool(args.sheetname is not None):
            print('Please provide both the cell and the sheetname, or none of the 2, not just one of them')
            sys.exit()
        elif args.cell is not None and args.sheetname is not None:
            starting_cell = Cell(args.sheetname, args.cell)
        else:
            print('Enter the sheetname and cell location of the cell you want to start the analysis from')
            sheetname = input("\tsheetname: ")
            cell_number = input("\tcell number: ")
            starting_cell = Cell(sheetname, cell_number)
        
        # TODO - Remove this hardcoded list of starting cells
        global nested_table_cells
        if cell_number in nested_table_cells:
            popped_value = nested_table_cells.pop(cell_number)
        global table_cells
        if cell_number in table_cells:
            table_cells.remove(cell_number)
        
        # Stack to keep track of formulas and values 
        formulas = Stack()
        values = Stack()
        exceptions = Stack()
        
        formulas, values, exceptions = resolve_cell(workbook, starting_cell, formulas, values, exceptions)
        print_results(formulas, values, exceptions, write_to_file=args.write_to_file)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--full_analysis', action='store_true', help='Run the full analysis, based on the hardcoded starting cells')
    parser.add_argument('--single_cell', action='store_true', help='Run the full analysis, based on the hardcoded starting cells')
    parser.add_argument('-f', '--file', type=str, help='The Excel file to extract the formulas from')
    parser.add_argument('-c', '--cell', type=str, help='The cell you want to start the analysis from')
    parser.add_argument('-sh', '--sheetname', type=str, help='The sheetname of the cell you want to start the analysis from')
    parser.add_argument('-wtf', '--write_to_file', action='store_true', help='Use this flag if you want to write the results to files')
    # TODO - add argument for the to ignore cells
    args = parser.parse_args()
    
    main(args)
