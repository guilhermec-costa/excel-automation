def extract_data(**kwargs):
    wb = kwargs['wb']
    new_row = kwargs['new_row']
    range = kwargs['range']
    for col in wb.sheet.range(range):
        column_letter = str(col.address)[1]
        first_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}20').end('right').row
        last_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}20').end('down').row
        if col.value in new_row.keys():
            new_row[col.value].extend \
            (wb.sheet.range(f'{column_letter}{first_row_index}:{column_letter}{last_row_index}').value)


def adjust_na_values(row):
    for key, value in row.items():
        for idx, item in enumerate(value):
            if item == None:
                value[idx] = 0

def eliminate_keys(row, keys=[]):
    for key in keys:
        del row[key]