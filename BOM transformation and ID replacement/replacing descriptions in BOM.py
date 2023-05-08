import xlwings as xw

# Make sure to save Excel workbook after running script 

wb = xw.Book('templates\Description replacement\S200seriesBOM-IMP_ForMiscItemList.xlsx')
ID_list = xw.Book('templates\Description replacement\S200 unique ALL.xlsx')

verified_sheet = ID_list.sheets['Verified'] # this sheet stores all the IDs, Eng Dash information, and DWG information

IDs_to_change = [511474.0, 513185.0, 513165.0, 511611.0, 501469.0, '501567 / 501568']
all_items_to_delete = [335525.0, 391215.0, 335527.0, 513428.0, 266712.0]


def find_verified_location(item_ID):

    ID_list = verified_range.value
    if item_ID not in ID_list:
        print("ID skipped: " + str(item_ID))
        return
    ID_location_row = ID_list.index(item_ID) + 2
    return ID_location_row



# argument is the row number of the item to update (descriptions, ID, etc.)
def make_changes(row_of_item, verified_location):

    sheet.range(f'M{row_of_item}').value = verified_sheet.range(f'C{verified_location}').value # replacing eng dash description
    sheet.range(f'D{row_of_item}').value = verified_sheet.range(f'D{verified_location}').value # replacing dwg descriptionv
    
    if sheet.range(f'L{row_of_item}').value in IDs_to_change:
        sheet.range(f'L{row_of_item}').value = verified_sheet.range(f'B{verified_location}').value
    print(f"Row values for {row_of_item}: {sheet.range(f'L{row_of_item}').value}, {sheet.range(f'M{row_of_item}').value}, {sheet.range(f'D{row_of_item}').value}")
    return



def checking_for_duplicate_deletion(value, row):

    if value in items_to_delete_per_sheet.values():
        print(f"Verification needed for drawing {sheet.name}: item {value} appears more than once; delete all entries?")
        deletion_conflict.append(value)
        return
    items_to_delete_per_sheet[row] = value



def deletion_cleanup():
    rows_to_remove = []
    for value_to_check in deletion_conflict:
        row_to_remove = [key for key, value in items_to_delete_per_sheet.items() if value == value_to_check]
        rows_to_remove = rows_to_remove + row_to_remove

    will_delete = [row for row in items_to_delete_per_sheet.keys() if row not in rows_to_remove] # removing all rows corresponding to IDs in deletion_conflict
    return will_delete



# list_of_rows will be populated with list of rows of IDs that should not be on the BOM
def delete_rows(list_of_rows):

    for row in reversed(list_of_rows): # deleting in reverse to not mess up row numbers
        sheet.range(f"A{row}").api.EntireRow.Delete()
    return



for sheet in wb.sheets:

    last_row = sheet.used_range.last_cell.row
    ID_range = sheet.range(f'L3:L{last_row}')
    verified_range = verified_sheet.range(f'A2:A{verified_sheet.used_range.last_cell.row}')

    print(sheet.name)

    items_to_delete_per_sheet = {} # dictionary of row:value pairs
    deletion_conflict = [] # list of IDs that occur in ID_to_delete_per_sheet but occur more than once in the sheet

    for cell in ID_range:

        value = cell.value
        if value in all_items_to_delete:
            checking_for_duplicate_deletion(cell.value, cell.row)

        verified_location = find_verified_location(value)
        if verified_location:
            make_changes(cell.row, verified_location)

    to_delete = deletion_cleanup()

    if to_delete:
        delete_rows(to_delete)
