from Util.Cell import Cell


def handle_max_min(cells, excel_max_min):
    max_min = excel_max_min[4:]
    arguments_unformatted = max_min.split(';')
    
    for arg in arguments_unformatted:
        # TODO handle recursion / nested functions
        cells.append(Cell('Tax Calculation', arg))
    
    return cells, excel_max_min