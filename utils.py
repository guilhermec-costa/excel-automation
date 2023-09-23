import xlwings as xw

class Workbook:
    def __init__(self, path):
        self.wb = xw.Book(path)
        self.sheet = None

    def go_to_sheet(self, sheet_name):
        self.sheet = self.wb.sheets(sheet_name)

    @staticmethod
    def create_empty_row():
        row = {
            'COD PRODUTO':[],
            'QTD DA OP':[],
            'QTD PROG.':[],
            'SALDO A PROG':[],
            'TOTAL SETUP': [],
            'META HORA TOP':[],
            'META HORA BOT':[],
            'TOTAL META HORA': [],
            'METADE META HORA': [],
            '3/4 TOTAL HORA': []
        }
        return row

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