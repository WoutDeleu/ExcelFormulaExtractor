class Cell:
    def __init__(self, sheetname, location):
        self.sheetname = sheetname
        self.location = location
        
    def __str__(self) -> str:
        return self.location
    
    
class CellValue:
    def __init__(self, cell, value):
        self.cell = cell
        self.value = value
        

class CellFormula:
    def __init__(self, cell, formula):
        self.cell = cell
        self.formula = formula