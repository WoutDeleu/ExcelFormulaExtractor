import unittest
from ExcelHandler.excel_extractor import extract_formula_cells


class TestExcelExtractor(unittest.TestCase):
    def test_max(self):
        cells, formula = extract_formula_cells('=MAX(A1:A3;B1:B3)')
        
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells:
            print(cell.location)
            self.assertTrue(cell.location in correct_cells)
        
        print(formula)
        self.assertTrue(formula == 'MAX(A1;A2;A3;B1;B2;B3)')
        
    def test_sum(self):
        cells, formula = extract_formula_cells('=SUM(A1:A3;B1:B3)')
        
        correct_cells = ['A1', 'A2', 'A3', 'B1', 'B2', 'B3']
        for cell in cells:
            print(cell.location)
            self.assertTrue(cell.location in correct_cells)
        
        print(formula)
        self.assertTrue(formula == 'A1+A2+A3+B1+B2+B3')

if __name__ == '__main__':
    unittest.main()