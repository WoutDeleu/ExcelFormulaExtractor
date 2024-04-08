import re


def remove_char_from_string(string, index):
    result_string = ''
    for i in range(len(string)):
        if i != index:
            result_string += string[i]
    return result_string

def is_int(var):
    try:
        int(var)
        return True
    except ValueError:
        return False

def is_float(var):
    try:
        float(var)
        return True
    except ValueError:
        return False
    
def list_to_string(list):
    if len(list) == 0:
        return '[]'
    result = '['
    for item in list:
        result += str(item) + ', '
    return result[:-2] + ']'

def is_letter_or_number(character):
    # Reguliere expressiepatroon om te controleren of iets een letter (hoofd- of kleine letter) of een getal is
    pattern = r'^[a-zA-Z0-9]$'
    return bool(re.match(pattern, character))

def is_percentage(string):
    # Reguliere expressiepatroon om te controleren of iets een percentage is
    pattern = r'^[0-9]+%$'
    return bool(re.match(pattern, string))

def format_namespace(namespace):
    return namespace.lower().replace(' ', '_')

def is_constant(string):
    return isinstance(string, int) or isinstance(string, float)

def write_results_to_file(data_structure, filename, name):
    if not data_structure.is_empty():
        file = open(filename+'_'+name+'.txt', 'w+')
        while not data_structure.is_empty():
            value = data_structure.pop()
            if name == 'formulas':
                file.write(format_namespace(value.cell.sheetname) + '_' + value.cell.location + '=' + str(value.formula) + '\n')
            else:
                file.write(format_namespace(value.cell.sheetname) + '_' + value.cell.location + '=' + str(value.value) + '\n')
        file.close()
    

def print_results(formulas, values, exceptions, to_file=True):
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
    print()
        
    print('######################################################################################################')
    print('#######################################  EXCEPTIONS  #################################################')
    print('######################################################################################################')
    for error in exceptions.get_list():
        # print(error.cell.location + ': ' + str(error.value))
        print(error.cell.location + '-' + error.cell.sheetname + ': ' + str(error.value))
    print()
    
    if to_file:
        filename = input('What is the filename where you want to write your results to? ')
        write_results_to_file(values, 'results/'+filename, 'values')
        write_results_to_file(formulas, 'results/'+filename, 'formulas')
        write_results_to_file(exceptions, 'results/'+filename, 'exceptions')