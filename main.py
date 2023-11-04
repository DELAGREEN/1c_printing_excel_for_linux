###!!!!!!!!minimum viable product

import sys
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
import time
import os


class textcolors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

param_list = f'''
{textcolors.YELLOW}-excel {textcolors.GREEN}активирует запись в EXCEL, принимает в качестве параметров {textcolors.BLUE}[путь к excel файлу][путь к txt файлу][разделитель1][разделитель2] 

'''

doc = f'''
{textcolors.YELLOW}-excel {textcolors.BLUE}[путь к excel файлу][путь к txt файлу][разделитель1][разделитель2]
        
{textcolors.YELLOW}-excel {textcolors.GREEN}флаг который показывает какого типа файл нужно будет собрать {textcolors.END}
{textcolors.YELLOW}[путь к excel файлу] {textcolors.GREEN} Путь куда конкретно сохранить файл Excel с расширением {textcolors.BLUE}(.xls|.xlsx|и др)
{textcolors.YELLOW}[путь к txt файлу] {textcolors.GREEN} Путь к txt файлу сформированный 1с в котором лежат матаданные сериализатора
{textcolors.YELLOW}[разделитель1] {textcolors.GREEN} Разделитель между данными по одной конкретной ячейки {textcolors.BLUE}[номер строки][разделитель1][номер столбца][разделитель1][значение ячейки]
{textcolors.YELLOW}[разделитель2] {textcolors.GREEN} Разделитель между данными ячеек {textcolors.BLUE}[номер строки][разделитель1][номер столбца][разделитель1][значение ячейки][разделитель2][номер строки][разделитель1][номер столбца][разделитель1][значение ячейки]

'''

def is_empty_excel_or_txt(_path_to_excel:str, _path_to_txt:str):
    '''Проверяем создан файл если нет создаем'''
    if not os.path.exists(_path_to_excel):
        wb = Workbook()                      
        wb.save(_path_to_excel)

    '''Проверяем создан файл если нет создаем'''
    if not os.path.exists(_path_to_txt):
        print(f'{textcolors.YELLOW}Ошибка. Файл с расширением .txt отсутствует')


def read_txt(_path_to_txt:str):
    with open(_path_to_txt, 'r') as txt_file:
        string = txt_file.read()
        
        #Вычищаем возможную помойку
        try:
            string = string.replace('\n', '')
        finally:
            try:
                string = string.replace('\ufeff', '')   #1С бывает засовывает свою какую то параметризацию выглядик как вот эта какаха которую я отрезаю
            finally:
                return string
        ##print(string)


'''
Функция принимает строку row:column:value,row:column:value,row:column:value 
дробит ее на параметры принимаемые openpyxl 
НАПРИМЕР:row:column:value для поячеечной записи в excel файл
'''
def writer_for_excel(_strings, _path_to_excel:str, _separator_first:str, _separator_end):
    wb = load_workbook(_path_to_excel)                              #Передаем классу, методу класса Файл который нужно открыть
    ws = wb.worksheets[0]                                           #Номер страницы для записи
    #дробим строку вида [номер строки, номер столбца, значение ячейки][разделитель][номер строки, номер столбца, значение ячейки]
    list_strings = _strings.split(_separator_end)                    #и получаем [номер строки][разделитель][номер столбца][разделитель][значение ячейки]                     

    for i in list_strings:
        _row, _column, _value = i.split(_separator_first)           #дробим строку на [номер строки][номер столбца][значение ячейки] 
        #ws[column] = value                                         #colunm должен равняться номеру столбца НАПРИМЕР: А1
        #print(_row, _column, _value)
        ws.cell(row=int(_row), column=int(_column)).value = _value   #передаем текст в ячейку                

    wb.save(_path_to_excel)                                          #сохраняем файл


def parse_param():

    param_name = sys.argv[1]
        
    if (param_name == '--excel' or param_name == '-excel'): #Если есть тригер выполняем код ниже
    
        try:
            path_to_excel = str(sys.argv[2])                         #Захватываем путь к excel
            path_to_txt = str(sys.argv[3])                           #Захватываем путь к txt
            separator_first = str(sys.argv[4])                       #Захватываем разделитель между значениями ячек и данными int[4]int[4]value
            separator_end = str(sys.argv[5])                         #Захватываем разделитель между данными одной ячейки int[4]int[4]value[5]int[4]int[4]value[5] и тд
            is_empty_excel_or_txt(path_to_excel, path_to_txt)
            writer_for_excel(read_txt(path_to_txt),path_to_excel, separator_first, separator_end)
            print(f'{textcolors.YELLOW}Complite: {param_name}')

        except Exception as _ex:
            if str(_ex) == 'list index out of range':
                print(f'{textcolors.RED}Ошибка.{textcolors.YELLOW} Недостаточное колличество параметров. Данный параметр принимает 4 дополнительных параметра. Обратитесь к -help')
            else: 
                print(f'{textcolors.RED}Ошибка.{textcolors.YELLOW}Неизвестная ошибка: Проверьте правильность параметра:{param_name}\n{_ex}')
            time.sleep(5)
            sys.exit(1)
        
    elif (param_name == '--help' or param_name == '-help'):
        print(f'{textcolors.CYAN}***HELP SHEET***\n{doc}')
        sys.exit(1)
   
    else:

        print(f'{textcolors.RED}Ошибка. {textcolors.YELLOW}Неизвестный параметр {param_name}')
        sys.exit(1)


if __name__ == '__main__':

    if len(sys.argv) == 1:

        print(f'{textcolors.RED}Внимание! {textcolors.YELLOW}Дополнительные параметры не заданы.\n' 
               'Модуль не работает без дополнительных параметров.\n'
               f'В качестве параметров принимается {textcolors.GREEN}Строка(String){textcolors.YELLOW} в качестве разделителей параметров {textcolors.GREEN}Пробел(Space).{textcolors.YELLOW} Пример: {textcolors.BLUE}[Параметр1] [Параметр2] [Параметр3]{textcolors.YELLOW} и тд (без квадратных скобок).\n'
               '\n'
               'На данный момент принимаются:'
               f'{param_list}\n'
               f'{textcolors.YELLOW}*ИЛИ*\n'
               'Ознакомтесь с документацией c помошью команды -help.'
               f'{textcolors.END}')
    
    elif len(sys.argv) == 2:
        parse_param()   #Парсим параметры

    elif len(sys.argv) == 6:
        parse_param()   #Парсим параметры 
        
    else:
        if len(sys.argv) <= 5:
            print(f'{textcolors.RED}Ошибка. {textcolors.YELLOW}Слишком мало параметров. Обратитесь к -help')
            sys.exit(1)

#time.sleep(5)
sys.exit(1)
