import xlwings as xw
import os
import colorama
from colorama import Fore, Back, Style
import time
colorama.init(autoreset=True)
import re

class Workbook:
    def __init__(self, path):
        self.wb = xw.Book(path)
        self.sheet = None
        self.existints_tabs = self.wb.sheet_names

    def go_to_sheet(self, sheet_name):
        self.sheet = self.wb.sheets(sheet_name)
        self.sheet.activate()

    def copy_sheet(self, sheet_name):
        self.sheet.copy(name=sheet_name)

    def close(self):
        self.wb.close()

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

def read_excel_path():
    while True:
        sheet_path = input(f"{Fore.YELLOW}Digite o caminho completo do arquivo > {Style.RESET_ALL}")
        if os.path.exists(sheet_path):
            print(f'{Fore.GREEN}Arquivo encontrado com sucesso!')
            return sheet_path
        print(f'{Fore.RED}Arquivo não encontrado. Digite novamente')
        print('----------------------------------')
        continue

def read_excel_tab(existing_tabs):
    default = "PROGRAMAÇÃO SMT L2"
    while True:
        sheet_name = str(input(f'{Fore.YELLOW}Digite o nome da aba no excel para o preenchimento{Style.RESET_ALL} {Fore.GREEN}(Padrão: {default}) > {Style.RESET_ALL}'))
        if sheet_name == "":
            return default
        if sheet_name in existing_tabs: 
            return sheet_name
        else:
            print(f'{Fore.RED}Verifique se a aba realmente existe na planilha especificada.')
            continue

def save_excel_tab(sheet:Workbook):
    default_location = sheet.wb.sheets("APOIO").range((4,5), (4,5)).value

    # fecha a nova aba aberta
    sheet.wb.app.quit()
    while True:
        path = input(f'{Fore.YELLOW}Digite o local de salvamento do arquivo{Style.RESET_ALL} {Fore.GREEN}(Padrão: {default_location}) > {Style.RESET_ALL}')
        if path == "":
            path = default_location
        if os.path.exists(path):
            return path
        else:
            print(f'{Fore.RED}Verifique se o caminho realmente existe no sistema.')
            continue

def display_title(title):
    print(Fore.CYAN)
    title_len = len(title)
    print(f'{Fore.CYAN}{"*"*title_len}')
    print(f'{Fore.WHITE}{title}')
    print(f'{Fore.CYAN}{"*"*title_len}')