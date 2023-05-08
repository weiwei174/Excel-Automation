import xlwings as xw

# adding variables for series name, units, and range to copy and extract values from for dynamism
# when a new series name is used, it only needs to be updated in one place (series_name)
# PLEASE ENSURE FORMULAS > CALCULATIONS IS SET TO AUTOMATIC IN EXCEL
# Occasional glitches in formatting can usually be fixed by interrupting, closing workbooks, and restarting

series_name = 'SSES-PYTHON'
units = 'IMP'

# opening required files
src_wb = xw.Book('templates\BOM transfomation\BOM template.xlsx')
working_wb = xw.Book(f'templates\BOM transfomation\{series_name} {units} template.xlsx')

# setting new path name for save as procedure, since a Save-As is required
dest_wb_path = f'./{series_name}seriesBOM-{units}_ForMiscItemList.xlsx'
working_wb.save(dest_wb_path)

# dest = new BOM to populate
dest_wb = xw.Book(dest_wb_path)

# src = template to pull from
src_sheet = src_wb.sheets['BOMTemplate']


# iterating through each sheet in  new wb:

for sheet in dest_wb.sheets:

    if sheet.name == 'Index':
        sheet.range('B5').value = f"3. See original {series_name} BOM for comments."
        continue
    
    starting_row = src_sheet.used_range.row                # stores first row of used range of template

    dest_BOM_last_row = sheet.used_range.last_cell.row     # stores last used row of BOM before any copy/paste

 
    # if the BOM is longer than 28 lines, insert more lines into the template to move everything down
    if dest_BOM_last_row > (starting_row - 6):
        src_sheet.range(f'20:{20 + dest_BOM_last_row - (starting_row - 6)}').insert('down')
        src_wb.save()


    range_to_copy = src_sheet.used_range.address

    src_sheet.range(range_to_copy).api.Copy()

    sheet.range(range_to_copy).api.PasteSpecial(-4104) # pastes with formatting


    # checking for invalid values
    for row in sheet.range(range_to_copy):
        for cell in row:
            if sheet.range(cell).value == "#REF!":
                    print(f"Error in {sheet.name} {cell.address}")


    row_to_copy = src_sheet.used_range.last_cell.row
    new_starting_row = src_sheet.used_range.row

    # checks if the BOM has more content than the template accounts for. if yes, this "extends" the template.
    if dest_BOM_last_row > 10:

        src_sheet.range(f'A{row_to_copy}:U{row_to_copy}').api.Copy()
        first_row_to_paste = row_to_copy + 1
        last_row_to_paste = row_to_copy + dest_BOM_last_row - 10
        print(first_row_to_paste, last_row_to_paste)
        sheet.range(f'A{first_row_to_paste}:U{last_row_to_paste}').api.PasteSpecial(-4104)
        range_to_extract_values = f'A{new_starting_row + 3}:R{last_row_to_paste}'

    else: 
        range_to_extract_values = f'A{starting_row + 3}:R{row_to_copy}'

    # sets this range of cells equal to the value of the cells rather than the formulas in the cell
    sheet.range(range_to_extract_values).value = sheet.range(range_to_extract_values).value


    # storing notes in column O of excel sheet in dictionary with WWID as key and note as value
    range_to_search_notes = f'O1:O{dest_BOM_last_row}'
    notes = {}
    for cell in sheet.range(range_to_search_notes):

        # Lines with just "Note :" is ignored, and 'Requires review' will be captured by general statement on index sheet
        if cell.value != None and "note" not in str(cell.value).lower() and "comment" not in str(cell.value).lower() and cell.value !='Requires review (see bluelines on drawing) ':
            note = cell.value
            row_number = cell.row
            id = sheet.range(row_number, 1).value      # stores WWID (first column of original data) as key
            if id == None:                             # accounting for notes that do not correspond to a WWID
                id = sheet.name
            notes[id] = note

    if notes:
        print(notes)    # allows verification and also tracks progress as the code goes through every sheet
        print(sheet.name)

    range_to_delete = sheet.range(f'1:{starting_row - 1}')
    sheet.range(range_to_delete).delete() # deleting original BOM

    # deleting rows with #REF! values
    rows_to_clean_up = []

    rows_in_new_BOM = sheet.used_range.last_cell.row
    for cell in sheet.range(f'A3:A{rows_in_new_BOM}'):
        if cell.value == None or cell.value == 0:
                rows_to_clean_up.append(cell.row)    # making list of rows to delete

    for row in reversed(rows_to_clean_up): # deleting in reverse to not mess up row numbers
        sheet.range(f"A{row}").api.EntireRow.Delete()

    # converts data into a table
    last_cell_after = sheet.used_range.last_cell.address
    range_to_make_table = sheet.range(f'A2:{last_cell_after}')
    table = sheet.tables.add(source=sheet.range(range_to_make_table).expand())
    table.show_headers = True
    table.style = 'Table Style Medium 2'

    sheet.autofit()

    # adds notes after the last used row of the sheet, skipping a line for every new note
    i = 2
    new_last_row = sheet.used_range.last_cell.row
    for key in notes:
        cell_for_note = f'A{new_last_row + i}'
        sheet.range(cell_for_note).value = f'Note for {key}: {notes[key]}'
        i+=2

    # to autoselect first cell of sheet
    sheet.range('A1').api.Copy()
    sheet.range('A1').api.PasteSpecial(-4104) 
    sheet.Zoom = 100


dest_wb.save()