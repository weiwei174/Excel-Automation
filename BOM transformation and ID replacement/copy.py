import xlwings as xw

# loading BOM to be edited (wb), sheet containing new bolt data (bolt_list) and sheet containing hardware data (hardware_list)
wb = xw.Book('templates\HDG replacement\S500 BEFORE HDG BOLT REPLACEMENT.xlsx')
bolt_list = xw.Book('templates\HDG replacement\HDG template.xlsx').sheets['Bolts']
hardware_list = xw.Book('templates\HDG replacement\HDG template.xlsx').sheets['Nuts and Washers']

# functionality is broken into helper functions, defined before the main loop

# searches in designated cells of excel table for nut ID
def search_for_nut_id(ID):
    nut_range = hardware_list.range(f'B2:C{last_row_hardware}')
    for cell in hardware_list.range(nut_range):
        if cell.value == ID:
            hardware_size = hardware_list.range(f'A{cell.row}').value
            ASTM_nut = hardware_list.range(f'B{cell.row}').value
            ISO_nut = hardware_list.range(f'C{cell.row}').value
            if ASTM_nut:
                return f'{int(ASTM_nut)} / {int(ISO_nut)}', hardware_size
            return f'{int(ISO_nut)}', hardware_size


# searches in designated cells of excel table for washer ID
def search_for_washer_id(ID):
    washer_range = f'D2:F{last_row_hardware}'
    for cell in hardware_list.range(washer_range):
        if cell.value == ID:
            hardware_size = hardware_list.range(f'A{cell.row}').value
            ASTM_washer = hardware_list.range(f'D{cell.row}').value
            ISO_washer = hardware_list.range(f'E{cell.row}').value
            if ASTM_washer:
                return f'{int(ASTM_washer)} / {int(ISO_washer)}', hardware_size
            return f'{int(ISO_washer)}', hardware_size


# replaces description for nut or washer 
def replace_nut_or_washer_description(cell, hardware_type, hardware_size):
    if hardware_type == 'nut':
        cell.value = f'NUT, HEAVY HEX, STRUCTURAL, {hardware_size}'
    if hardware_type == 'washer':
        cell.value = f'FLAT WASHER, STRUCTURAL, {hardware_size}'


# searches every cell in 3 columns to check if current ID is a bolt, returns the row (containing all IDs with matching size and length) of bolt
def search_for_bolt_id(list_of_col, ID):
    for col in list_of_col:
        for cell in bolt_list.range(f'{col}2:{col}{last_row_bolt}'):
            if cell.value == ID:
                # print(f'cell row: {cell.row}')
                return cell.row
    
    return 0


# checks across bolt table for corresponding IDs in same length and size
def find_other_bolt_ids(row):
    ID_dict = {}
    for cell in bolt_list.range(f'C{row}:E{row}'):
        if cell.value: ID_dict[cell.column] = cell.value
    return ID_dict


# creates new ID based on various criteria and availability and determines whether a comment is required 
def replace_bolt_ids(dict, comment):
    ID = ''
    ISO = ''
    print(f'dict keys: {dict.keys()}')
    if 3 in dict.keys():
        ASTM = int(dict[3])
        ID = str(ASTM)
    
    if ID:
        if 5 in dict.keys():
            ISO = str(int(dict[5]))
            ID += f' / {ISO}'
        elif 4 in dict.keys():
            ISO = str(int(dict[4]))
            ID += f' / {ISO}'
            comment = ISO
    else:
        if 5 in dict.keys():
            ID = str(int(dict[5]))
            print(f'ISO PT: {ID}')
        elif 4 in dict.keys():
            ID = str(int(dict[4]))
            comment = ID
            # print(f'comment in function: {comment}')
    print(f'ID: {ID}')
    return ID, comment


# replaces description for bolt
def replace_bolt_description(cell):
    size = bolt_list.range(f'A{wanted_row}').value
    length = int(bolt_list.range(f'B{wanted_row}').value)
    cell.value = f'BOLT, HEAVY HEX HEAD, STRUCTURAL, {size} X {length}'
    # print(cell.value)


# prints comment if fully threaded bolt
def comment_handler(comment_id, item_row):
    target_cell = sheet.range(f'B{last_row + 1}')
    item_callout = sheet.range(f'C{item_row}').value
    target_cell.value = f'Item {item_callout}, ID {comment_id}, is fully threaded.'
    target_cell.api.HorizontalAlignment = -4131




# MAIN LOOP

for sheet in wb.sheets:
    if sheet.name != 'S500-4':
        break

    last_row = sheet.used_range.last_cell.row # calculates last used row of current BOM to enable adding notes to empty cell at end of sheet
    ID_range = sheet.range(f'L3:L{last_row}') # calculates last used row of bolt sheet to enable easier search of IDs
    last_row_bolt = bolt_list.used_range.last_cell.row # calculates last used row of bolt sheet to enable easier search of IDs
    last_row_hardware = hardware_list.used_range.last_cell.row

    print(sheet.name) # to trouble shoot which BOM is causing an issue
    for cell in ID_range:
        comment = '' # resets comment for every new ID
        description_cell = sheet.range(f'M{cell.row}')

        if not cell.value:
            continue

        if type(cell.value) == str:
            if '/' in cell.value:
                continue
        
        if ('BOLT' not in description_cell.value) and ('SCREW' not in description_cell.value) and ('NUT' not in description_cell.value) and ('WASHER' not in description_cell.value):
            continue
 
        elif ('BOLT' not in description_cell.value) and ('SCREW' not in description_cell.value) and ('NUT' in description_cell.value):
            nut_id = search_for_nut_id(cell.value)
            if not nut_id:
                # print('ERROR: no nut found')
                continue
            cell.value = nut_id[0]                  
            nut_size = nut_id[1]
            replace_nut_or_washer_description(description_cell, 'nut', nut_size) # replacing hardware description in BOM
            continue

        elif ('WASHER' in description_cell.value):
            washer_id = search_for_washer_id(cell.value)
            if not washer_id:
                # print('ERROR: no washer found')
                continue
            cell.value = washer_id[0]                  
            washer_size = washer_id[1]
            replace_nut_or_washer_description(description_cell, 'washer', washer_size) # replacing hardware description in BOM
            continue
        
        if 'HEAVY' not in description_cell.value and 'HEXAGON' not in description_cell.value:
            print('Not structural bolt, skipped.')
            continue
        
        wanted_row = search_for_bolt_id(['C', 'D', 'E'], cell.value) # searching ID columns of bolt template for the current ID 
        if not wanted_row:
            print('ERROR: no bolt found')
            continue
        ID_dict = find_other_bolt_ids(wanted_row) # looks for corresponding IDs in the same size and stores the available IDs by type in a dictionary
        new_id = replace_bolt_ids(ID_dict, comment) # replaces ID of bolt(s) based on specific criteria of ASTM/ISO if both are available
        cell.value = new_id[0]
        comment = new_id[1]
        replace_bolt_description(description_cell) # replaces description of bolts based on size and length

        if comment:
            comment_handler(comment, cell.row)
            last_row = sheet.used_range.last_cell.row
            print('Comment on ', sheet.name)

# wb.save('S500 HDG replaced')