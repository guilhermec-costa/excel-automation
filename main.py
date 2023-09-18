import xlwings as xw
import math

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


wb = Workbook('FM613programacao_SMT_OUT_2023 _TESTE.xlsb')
wb.go_to_sheet('PROGRAMAÇÃO SMT')
new_row = wb.create_empty_row()

# extração de dados da planilha
for col in wb.sheet.range('E10:R10'):
    column_letter = str(col.address)[1]
    first_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}200').end('left').row
    last_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}200').end('down').row
    if col.value in new_row.keys():
        new_row[col.value].extend \
        (wb.sheet.range(f'{column_letter}{first_row_index}:{column_letter}{last_row_index}').value)


# construção da lista de valores para formação de setup
for idx, hour_goal in enumerate(new_row['META HORA TOP']):
    if hour_goal > 0:
        new_row['TOTAL META HORA'].append(hour_goal)
    else:
        new_row['TOTAL META HORA'].append(new_row['META HORA BOT'][idx])

# substituição de valores vazios por zeros
for key, value in new_row.items():
    for idx, item in enumerate(value):
        if item == None:
            value[idx] = 0

# arredondamento das horas de setup para hora superior mais próxima
new_row['TOTAL SETUP'] = [math.ceil(setup) for setup in new_row['TOTAL SETUP']]
del new_row['META HORA TOP']
del new_row['META HORA BOT']

# construção das listas de meta hora, meta 50% hora e meta 75% hora
new_row['TOTAL META HORA'] = [math.trunc(hour_goal) for hour_goal in new_row['TOTAL META HORA']]
new_row['METADE META HORA'] = [total_goal_hour/2 for total_goal_hour in new_row['TOTAL META HORA']]
new_row['3/4 TOTAL HORA'] = [math.trunc(total_goal_hour * 0.75) for total_goal_hour in new_row['TOTAL META HORA']]

for key, value in new_row.items():
    print(key)
    print(value)
    print('Size: ', len(value))
    print('')
print('--------------------------------------------')

# posicionamento de setups
for col in wb.sheet.range('E10:R10'):
    column_letter = str(col.address)[1]
    first_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}200').end('left').row
    last_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}200').end('down').row
    if col.value == 'COD PRODUTO':
        column_letter = str(col.address)[1]
        for idx, row in enumerate(wb.sheet.range\
                            (f'{column_letter}{first_row_index}:{column_letter}{last_row_index}')):
            op = row.value
            op_line = row.row
            op_position = column_letter + str(op_line)
            setup_for_op = new_row['TOTAL SETUP'][idx]

            start_col_position = 21
            cols = ['VWXYZ']
            for setup in range(setup_for_op):
                print(op_line, start_col_position)
                wb.sheet[op_line-1, start_col_position].value = 'setup'
                # print(f'{op_line}:{op_line+1}, {start_col_position}:{start_col_position+1}')
                # print(start_col_position, op_line)
                start_col_position += 1