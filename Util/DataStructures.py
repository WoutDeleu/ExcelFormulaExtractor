from Util.Cell import Cell, CellFormula, CellValue


class Queue:
    def __init__(self):
        self.queue = []
        
    def add(self, item):
        self.queue.append(item)
        
    def pop(self):
        return self.queue.pop(0)
    
    def is_empty(self):
        return len(self.queue) == 0
    
    def size(self):
        return len(self.queue)
    
    def get_list(self):
        return self.queue
    
    def __str__(self):
        return str(self.queue)


class Stack:
    def __init__(self):
        self.stack = []
        
    def add(self, cell):
        if not self.contains(cell):
            self.stack.append(cell)
        else:
            index = self.find_item_index(cell)
            self.stack.pop(index)
            self.stack.append(cell)
            
    def move_to_top(self, item):
        index = self.find_item_index(item)
        item = self.stack.pop(index)
        self.stack.append(item)
            
    def find_item_index(self, item):
        if isinstance(item, Cell):
            for i in range(len(self.stack)):
                if self.stack[i].cell.location == item.location and self.stack[i].cell.sheetname == item.sheetname:
                    return i
        elif isinstance(item, CellFormula) or isinstance(item, CellValue):
            for i in range(len(self.stack)):
                if self.stack[i].cell.location == item.cell.location and self.stack[i].cell.sheetname == item.cell.sheetname:
                    return i
        
    def contains(self, item):
        if isinstance(item, Cell):
            for cell in self.stack:
                if cell.cell.location == item.location and cell.cell.sheetname == item.sheetname:
                    return True
            return False
        elif isinstance(item, CellFormula) or isinstance(item, CellValue):
            for cell in self.stack:
                if cell.cell.location == item.cell.location and cell.cell.sheetname == item.cell.sheetname:
                    return True
            return False
        else:
            return False
    
    def contains_cell(self, item):
        for cell in self.stack:
            if cell.cell.location == item.cell.location and cell.cell.sheetname == item.cell.sheetname:
                return True
        return False
        
    def pop(self):
        return self.stack.pop()
    
    def is_empty(self):
        return len(self.stack) == 0
    
    def size(self):
        return len(self.stack)
    
    def get_list(self):
        return self.stack
    
    def __str__(self):
        return str(self.stack)


class Set:
    def __init__(self):
        self.set = []
        
    def append(self, item):
        if not self.contains(item):
            self.set.append(item)
            
    def get_list(self):
        return self.set
    
    def contains(self, item):
        if isinstance(item, Cell):
            for cell in self.set:
                if cell.location == item.location and cell.sheetname == item.sheetname:
                    return True
            return False
    
    def __str__(self):
        return str(self.set)