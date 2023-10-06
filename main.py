import time
import shutil
import math
import utils
import os
import datetime
from tqdm import tqdm
import xlwings as xl

DATE_FORMATTED = datetime.datetime.now().strftime('%d-%m-%Y-%H-%M-%S')
NAME = "FM-613 Programação SMT"
OPS_EXCEPTIONS = ('MANUT', 'PPROG', "manut", "pprog", "OPANT", "opant")
# inicialização do app

title = 'Programação SMT'
utils.display_title(title)
sheet_path = utils.read_excel_path()
tmp_sheet = utils.Workbook(sheet_path)
print()
path_to_save = utils.save_excel_tab(tmp_sheet)
path_to_save_final_result = path_to_save + f"{NAME}_{DATE_FORMATTED}.xlsb"

shutil.copyfile(fr"{sheet_path}", fr"{path_to_save_final_result}")
print()
print('Cópia do arquivo inicial criada em: ', path_to_save_final_result)

wb = utils.Workbook(fr"{path_to_save_final_result}")
wb.go_to_sheet(utils.read_excel_tab(wb.existints_tabs))
new_row = wb.create_empty_row()

new_row['1/4 TOTAL HORA'] = []
new_row['3/4 TOTAL HORA'] = []
new_row['METADE META HORA']= []

# extração de dados da planilha
print()
print('Extraindo dados...')
utils.extract_data(wb=wb, new_row=new_row, range='E10:R10')
time.sleep(1.5)

# construção da lista de valores para formação de setup
print()
print('Ajustando valores para posicionamento dos setups')
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
new_row['TOTAL META HORA'] = [math.floor(hour_goal) for hour_goal in new_row['TOTAL META HORA']]

# correção de casos em que a meta hora é menor que a qtd da OP
for value in new_row['TOTAL META HORA']:
    new_row['1/4 TOTAL HORA'].append(math.floor(value * 0.25))
    new_row['3/4 TOTAL HORA'].append(math.floor(value * 0.75))
    new_row['METADE META HORA'].append(math.floor(value/2))

print()
print('Posicionando setups e distribuindo peças...')
start_col_position = 21
for col in wb.sheet.range((10, 5), (10, 18)):
    first_row_index = wb.sheet.range((11, col.column), (20, col.column)).end('left').row
    last_row_index = wb.sheet.range((11, col.column), (20, col.column)).end('down').row

    if col.value == 'COD PRODUTO':
        col_values = enumerate(wb.sheet.range((first_row_index, col.column), (last_row_index, col.column)))
        pbar = tqdm(total=last_row_index - first_row_index+1)
        for idx, row in tqdm(col_values):
            time.sleep(0.20)
            pbar.update(1)
            op_value = row.value
            op_line = row.row-1
            setup_for_op = new_row['TOTAL SETUP'][idx]
            quarter_goal_hour = new_row['1/4 TOTAL HORA'][idx]
            three_quarters_goal = new_row['3/4 TOTAL HORA'][idx]
            half_goal = new_row['METADE META HORA'][idx]
            full_goal = new_row['TOTAL META HORA'][idx]
            total_qtd_op = new_row['QTD DA OP'][idx]
            print("Full goal: ", full_goal, "        ")
            
            counter_setup, counter_setup_extra, total_goal = 0, 0, 0

            # print('Posicionado setups')
            while counter_setup < setup_for_op:
                if wb.sheet[2, start_col_position].value in ('SIM', 'U'):
                    if op_value in OPS_EXCEPTIONS:
                        wb.sheet[op_line, start_col_position].value = str(op_value).lower()
                        wb.sheet[op_line, start_col_position].autofit()
                    else:
                        wb.sheet[op_line, start_col_position].value = 'setup'


                    counter_setup += 1
                start_col_position += 1

            # para posição do setup anterior
            # if op_line > 10:
            #     wb.sheet[op_line, start_col_position-setup_for_op-1].value = 'setup'
            

            counter_first_setup = 1
            #last_hour_of_work = 18
            # posicionamento de qtd
            while True:
                current_work_hour_object = wb.sheet[8, start_col_position]
                if wb.sheet[2, start_col_position].value == 'U':
                    last_hour_of_work = int(current_work_hour_object.value)
                    if wb.sheet[8, start_col_position].value == last_hour_of_work:
                        if counter_first_setup == 1:
                            value_to_add = quarter_goal_hour
                        elif counter_first_setup == 2:
                            value_to_add = half_goal
                        else:
                            value_to_add = three_quarters_goal
                elif wb.sheet[2, start_col_position].value == 'SIM':
                    if full_goal == 0 or total_qtd_op == 0:
                        start_col_position -= 1
                        break
                    if counter_first_setup == 1:
                        value_to_add = quarter_goal_hour
                    elif counter_first_setup == 2:
                        value_to_add = half_goal
                    else:
                        if wb.sheet[8, start_col_position].value == 8:
                            value_to_add = three_quarters_goal
                        else:
                            value_to_add = full_goal
                if wb.sheet[2, start_col_position].value in ('SIM', 'U'):
                    total_goal += value_to_add
                    if total_goal <= total_qtd_op:
                        wb.sheet[op_line, start_col_position].value = value_to_add
                        # wb.sheet.range((1, start_col_position), (last_row_index, start_col_position)).columns.autofit()
                    else:
                        wb.sheet[op_line, start_col_position].value = value_to_add
                        if total_goal > total_qtd_op:
                            diff = total_goal - total_qtd_op

                            wb.sheet[op_line, start_col_position].value -= diff
                            # wb.sheet.range((1, start_col_position), (last_row_index, start_col_position)).columns.autofit()
                        break
                    counter_first_setup += 1
                start_col_position += 1
            start_col_position += 1
pbar.close()

# adição de colunas de total
# print()
# print('Posicioando colunas de total...')
# start_col = 22
# last_hour_column_index = wb.sheet.range((9, start_col), (9, 500)).end('right')
# for hour in wb.sheet.range((9, start_col), (9, last_hour_column_index.column)):
#     hour_value = hour.value
#     hour_column = hour.column
#     if not hour_value in (None, 'Total'):
#         if int(hour_value) == 3:
#             column_text = wb.sheet[9:9, hour_column:hour_column].address[1:3]
#             wb.sheet.range(f'{column_text}:{column_text}').insert('down')
#             wb.sheet[8, hour_column].value = 'Total'
#             wb.sheet[8, hour_column].font.bold = True
#             wb.sheet[8, hour_column].color = (127, 235, 250)
#             wb.sheet[9, hour_column].value = '-'

#     elif hour_value == 'Total':
#         range_to_sum = wb.sheet.range((11, hour_column-1), (11, hour_column-20))
#         total_to_sum = sum([value for value in range_to_sum.value if isinstance(value, (int, float))])
#         addres_to_sum = str(range_to_sum.address).replace('$', "")
#         wb.sheet[10, hour.column-1].formula = f"=SUM({addres_to_sum})"
#         first_row_index = wb.sheet.range(addres_to_sum).end('right').row
#         last_row_index = wb.sheet.range(addres_to_sum).end('down').row
        
#         wb.sheet.range((first_row_index, hour_column), (last_row_index, hour_column)).clear_formats()
#         wb.sheet.range((first_row_index, hour_column), (last_row_index, hour_column)).font.bold = True
#         wb.sheet.range((first_row_index, hour_column), (last_row_index, hour_column)).color = (255, 166, 43)

if input('Pressione enter para salvar o arquivo') == "":
    wb.wb.save()
    print(f'Arquivo salvo em {path_to_save_final_result}')
if input('Pressione enter para fechar o programa') == "":
    print('Finalizando programa em 3 segundos...')
    time.sleep(3)
    exit()
    