import xlwings as xw
import math
import utils

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
wb.go_to_sheet('PROGRAMAÇÃO_SMT_teste')
new_row = wb.create_empty_row()
# extração de dados da planilha
print('extração de dados da planilha')
utils.extract_data(wb=wb, new_row=new_row, range='E10:R10')

print('--------------------------')
# construção da lista de valores para formação de setup
for idx, hour_goal in enumerate(new_row['META HORA TOP']):
    if hour_goal > 0:
        new_row['TOTAL META HORA'].append(hour_goal)
    else:
        new_row['TOTAL META HORA'].append(new_row['META HORA BOT'][idx])

# substituição de valores vazios por zeros
utils.adjust_na_values(new_row)

# arredondamento das horas de setup para hora superior mais próxima
new_row['TOTAL SETUP'] = [math.ceil(setup) for setup in new_row['TOTAL SETUP']]
utils.eliminate_keys(['META HORA TOP', 'META HORA BOT'])

# construção das listas de meta hora, meta 50% hora e meta 75% hora
new_row['TOTAL META HORA'] = [math.trunc(hour_goal) for hour_goal in new_row['TOTAL META HORA']]


# correção de casos em que a meta hora é menor que a qtd da OP
new_row['1/4 TOTAL HORA'] = []
new_row['3/4 TOTAL HORA'] = []
new_row['METADE META HORA']= []
for idx, item in enumerate(new_row['TOTAL META HORA']):
    item_on_qtd_op = new_row['QTD DA OP'][idx]
    if item_on_qtd_op >= item:
        new_row['1/4 TOTAL HORA'].append(math.trunc(item * 0.25))
        new_row['3/4 TOTAL HORA'].append(math.trunc(item * 0.75))
        new_row['METADE META HORA'].append(math.trunc(item/2))
    else:
         new_row['1/4 TOTAL HORA'].append(math.trunc(item_on_qtd_op * 0.25))
         new_row['3/4 TOTAL HORA'].append(math.trunc(item_on_qtd_op * 0.75))
         new_row['METADE META HORA'].append(item_on_qtd_op/2)

for key, value in new_row.items():
    print(key)
    print(value)
    print('Size: ', len(value))
    print('')
print('--------------------------------------------')
# posicionamento de setups

print('Posicionado setups')
for col in wb.sheet.range('E10:R10'):
    column_letter = str(col.address)[1]
    first_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}20').end('left').row
    last_row_index = wb.sheet.range(f'{column_letter}11:{column_letter}20').end('down').row
    if col.value == 'COD PRODUTO':
        column_letter = str(col.address)[1]
        start_col_position = 21
        for idx, row in enumerate(wb.sheet.range\
                            (f'{column_letter}{first_row_index}:{column_letter}{last_row_index}')):
 

            op_line = row.row-1
            op_position = column_letter + str(op_line)
            setup_for_op = new_row['TOTAL SETUP'][idx]
            quarter_goal_hour = new_row['1/4 TOTAL HORA'][idx]
            three_quarters_goal = new_row['3/4 TOTAL HORA'][idx]
            half_goal = new_row['METADE META HORA'][idx]
            full_goal = new_row['TOTAL META HORA'][idx]
            total_qtd_op = new_row['QTD DA OP'][idx]
            
            counter_setup, counter_setup_extra, total_goal = 0, 0, 0
            while counter_setup < setup_for_op:
                if wb.sheet[2, start_col_position].value == None:
                    wb.sheet[op_line, start_col_position].value = 'setup'
                    counter_setup += 1
                start_col_position += 1

            # if op_line > 10:
            #     wb.sheet[op_line, start_col_position-setup_for_op-1].value = 'setup'
            
            counter_first_setup = 1
            # posicionamento de qtd
            while True:
                if full_goal == 0:
                    break
                if wb.sheet[2, start_col_position].value == None:
                    print('Endereço:', wb.sheet[op_line, start_col_position].address)
                    if counter_first_setup == 1:
                        value_to_add = quarter_goal_hour
                    elif counter_first_setup == 2:
                        value_to_add = half_goal
                    else:
                        # pegar lista de X. Se wb.sheet[8, start_col_position].value for igual a coluna (último x - 1) na linha 8, então 75
                        # last_work_hour = 
                        if int(wb.sheet[8, start_col_position].value) in (8, 17, 20):
                            print('Horário', wb.sheet[8, start_col_position].value)
                            value_to_add = three_quarters_goal
                        else:
                            value_to_add = full_goal

                    total_goal += value_to_add
                    if total_goal <= total_qtd_op:
                        wb.sheet[op_line, start_col_position].value = value_to_add
                    else:
                        wb.sheet[op_line, start_col_position].value = value_to_add
                        if total_goal > total_qtd_op:
                            diff = total_goal - total_qtd_op

                            # corrige a diferença
                            wb.sheet[op_line, start_col_position].value -= diff
                        break
                    counter_first_setup += 1
                start_col_position += 1
            start_col_position += 1
            #caso a linha seja depois da 11, preenche com "setup" a célula abaixo do último setup da linha anterior
            print('Fora do loop')