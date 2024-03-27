import unittest
from ExcelHandler.excel_extractor import extract_formula_cells


class TestExcelExtractor(unittest.TestCase):
    def test_max(self):
        cells, formula = extract_formula_cells('=MAX(A1:A3;B1:B3)')
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == 'MAX(A1;A2;A3;B1;B2;B3)')
        
        
    def test_sum(self):
        cells, formula = extract_formula_cells('=SUM(A1:A3;B1:B3;A1)')
        
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == '(A1+A2+A3+B1+B2+B3+A1)')
        
    def test_if(self):
        cells, formula = extract_formula_cells('=IF(A1>1;A1;B1)')
        
        correct_cells = ['A1', 'B1']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == "IF(A1>1)'{'A1'}'else'{'B1'}'")    
    
    def test_sum_max_min_combined(self):
        cells, formula = extract_formula_cells('=SUM(A1:A3;B1:B3;MAX(A1;B1))+A1+MIN(A1:A3;B1:B3)')
        
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells.get_list():
            self.assertTrue(cell.location in correct_cells)
        
        self.assertTrue(formula == '(A1+A2+A3+B1+B2+B3+MAX(A1;B1))+A1+MIN(A1;A2;A3;B1;B2;B3)')

if __name__ == '__main__':
    unittest.main()