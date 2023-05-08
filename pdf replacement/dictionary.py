import xlwings as xw

# loading BOM to be edited (wb), sheet containing new bolt data (bolt_list) and sheet containing hardware data (hardware_list)
series = 'S500'
old_bom = xw.Book(f'templates\HDG replacement\{series} BEFORE HDG BOLT REPLACEMENT.xlsx') # BOM containing old IDs in drawings to be replaced
new_bom = xw.Book(f'templates\HDG replacement\{series}seriesBOM-MET_ForMiscItemList.xlsx') # BOM containing new IDs to override old IDs
bolt_list = xw.Book('templates\HDG replacement\HDG template.xlsx').sheets['Bolts'] # contains list of ASTM, ISO FT, and ISO PT fasteners
bolt_last_row = bolt_list.used_range.last_cell.row

# dictionary mapping old IDs as keys to new IDs as values
old_to_new = {}

# list of sheet names to enable PDFs to be dynamically loaded
sheet_names = []

for sheet in old_bom.sheets:
    name = sheet.name
    if name == 'Note':
        continue
    sheet_names.append(name)

    old_sheet = sheet
    new_sheet = new_bom.sheets[name]
    last_row = old_sheet.used_range.last_cell.row

    # IDs in BOM are stored in the L column; these return a list of the values in the L column
    old_id_list = old_sheet.range(f'L3:L{last_row}').value
    new_id_list = new_sheet.range(f'L3:L{last_row}').value


    for id in old_id_list:
        
        if id is None or type(id) == str or not id or id in old_to_new.keys():
            continue

        old_id = str(int(id)).strip()
        location = old_id_list.index(id)
        new_id = new_id_list[location]

        if type(new_id) == float:
            new_id = str(int(new_id)).strip()

        if new_id == old_id:
            continue

        ft = False
        for bolt in bolt_list.range(F'D2:D{bolt_last_row}').value:

            if not bolt:
                continue

            bolt = str(int(bolt)).strip()

            if bolt in new_id:
                # print(bolt)
                old_to_new[old_id] = f'{new_id} (FT)'
                # print(old_to_new[old_id])
                ft = True
    
        if ft == False:
            old_to_new[old_id] = new_id


print(old_to_new)
print(sheet_names)
