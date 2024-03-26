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
    
def list_to_string(list):
    result = '['
    for item in list:
        result += str(item) + ', '
    return result[:-2] + ']'

def is_letter_or_number(character):
    # Reguliere expressiepatroon om te controleren of iets een letter (hoofd- of kleine letter) of een getal is
    pattern = r'^[a-zA-Z0-9]$'
    return bool(re.match(pattern, character))
