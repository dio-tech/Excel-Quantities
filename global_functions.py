from multiprocessing import current_process
import string

def add_quantity(item_name, quantities, new_quantity, index):
    if not ' ' in str(new_quantity):
        for i in range(len(str(new_quantity))-1):
            if str(new_quantity)[i] + str(new_quantity)[i+1] == 'm2':
                new_quantity = str(new_quantity).replace(str(new_quantity)[i], " " + str(new_quantity)[i])

    k = [char for char in str(new_quantity)]
    digits = []
    string = []
    for i in range(len(k)):
        if k[i] == "x":
            k[i] = "*"
        
        if k[i].isdigit() or k[i] == "," or k[i] == "." or k[i] == "*":
            digits.append(k[i])
        else:
            string.append(k[i])
    if not ' ' in string:
        new_quantity = str(eval(''.join(digits))).replace('(', '').replace(')', '').replace(' ', '') + ' ' + str(''.join(string))    

    if quantities[index] != 0:
        if type(quantities[index]) != int and type(quantities[index]) != float:
            quantity = quantities[index].split(' ')
            if quantity[1] == '' or quantity[1] == ' ':
                quantity[1] = 'uni'
            if str(new_quantity).split(' ')[1] == '':
                quantity[0] = check_units(quantity[0], quantity[1], 'uni', item_name)
            else:
                quantity[0] = check_units(quantity[0], quantity[1], str(new_quantity).split(' ')[1], item_name)
            return str(float(quantity[0]) + float(new_quantity.split(' ')[0].replace(',', '.'))) + ' ' + quantity[1]
        else:
            return float(quantities[index]) + float(new_quantity)
    else:
        quantity = new_quantity.split(' ')
        if quantity[1] == '' or quantity[1] == ' ':
            quantity[1] = 'uni'
            return new_quantity.replace(',', '.') + quantity[1]
        return new_quantity.replace(',', '.')

def get_sop(file):
    f = file.split('_')
    for i in range(len(f)):
        if f[i].startswith('SOP'):
            return f[i]
    return f[2]

def get_rows(sheet):
    materials_row = None
    tools_row = None
    quantities_col = None
    for row in range(1, sheet.max_row):
        if sheet.cell(row, 1).value == "Materiais /Materials":
            materials_row = row
            break
    
    for row in range(1, sheet.max_row):
        if sheet.cell(row, 1).value == "Ferramentas/ Tools":
            tools_row = row
            break
    for col in range(1, sheet.max_column):
        if sheet.cell(materials_row, col).value == "Quant. [units]":
            quantities_col = col
            break
    
    return materials_row, tools_row, quantities_col

def check_units(new_quantity, current_unit, new_unit, item_name):
    if str(current_unit) != str(new_unit):
        print(f"Unidades de {item_name} são diferentes das atuais (atualmente: {current_unit})")
        quantity = input(f"Inserir quantidade total de {item_name} em {current_unit} (No ficheiro: {new_quantity} {new_unit}): ")
    
        return float(quantity)
    return float(new_quantity)
