import unittest
import warnings
from ExcelHandler.excel_extractor import extract_formula_cells
from ExcelHandler.excel_helpers import read_in_excel, split_up_conditions
from Util.DataStructures import Set

warnings.simplefilter(action='ignore', category=UserWarning)
workbook = read_in_excel('Draft PB-berekening - WERKVERSIE V4.xlsx')

class TestExcelExtractor(unittest.TestCase):
    
    def test_max(self):
        cells, formula = extract_formula_cells('Tax Calculation', '=MAX(A1:A3;B1:B3)', cells=Set())
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == 'MAX(tax_calculation_A1;tax_calculation_A2;tax_calculation_A3;tax_calculation_B1;tax_calculation_B2;tax_calculation_B3)')
        
        
    def test_sum(self):
        cells, formula = extract_formula_cells('Tax Calculation', '=SUM(A1:A3;B1:B3;A1)', cells=Set())
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == '(tax_calculation_A1+tax_calculation_A2+tax_calculation_A3+tax_calculation_B1+tax_calculation_B2+tax_calculation_B3+tax_calculation_A1)')
        
    def test_sum_max_min_combined(self):
        cells, formula = extract_formula_cells('Tax Calculation', '=SUM(A1:A3;B1:B3;MAX(A1;B1))+A1+MIN(A1:A3;B1:B3)', cells=Set())
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == '(tax_calculation_A1+tax_calculation_A2+tax_calculation_A3+tax_calculation_B1+tax_calculation_B2+tax_calculation_B3+MAX(tax_calculation_A1;tax_calculation_B1))+tax_calculation_A1+MIN(tax_calculation_A1;tax_calculation_A2;tax_calculation_A3;tax_calculation_B1;tax_calculation_B2;tax_calculation_B3)')

    def test_if(self):
        cells, formula = extract_formula_cells('Tax Calculation', '=IF(A1>1;A1;B1)', cells=Set())
        correct_cells = ['A1', 'B1']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == "IF(tax_calculation_A1>1){tax_calculation_A1}else{tax_calculation_B1}")    
        
    def test_if_and_or(self):
        cells, formula = extract_formula_cells('Tax Calculation', '=IF(AND(OR(A1>1;B1>1);(A2<Q1));A1;B3)', cells=Set())
        correct_cells = ['A1', 'B1', 'A2', 'B3', 'Q1']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        self.assertTrue(formula == "IF(((tax_calculation_A1>1|tax_calculation_B1>1)&tax_calculation_A2<tax_calculation_Q1)){tax_calculation_A1}else{tax_calculation_B3}")
        
    def test_if_not(self):
        cells, formula = extract_formula_cells('Tax Calculation', '=IF(NOT(A1>1);A2;B2)', cells=Set())
        correct_cells = ['A1', 'B1', 'A2', "B2"]
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        self.assertTrue(formula == "IF(!(tax_calculation_A1>1)){tax_calculation_A2}else{tax_calculation_B2}")
    
    def test_split_up_conditions(self):
        parts, operators = split_up_conditions('A1>1') 
        self.assertTrue(parts == ['A1','1'])
        self.assertTrue(operators == ['>'])
    
if __name__ == '__main__':
    unittest.main()
