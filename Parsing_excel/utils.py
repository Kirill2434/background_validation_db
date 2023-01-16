import os

import glob
from pprint import pprint

import openpyxl
import pandas as pd
from pathlib import Path

from pandas import ExcelWriter
from tqdm import tqdm

from source.config import COL_NAME, SHEET_NAME_122, FILE_NAME_CHECK_REPORT, main_folder, files_source
from source.custom_exceptions import EmptyException


def check_directory():
    """Проверка директории "С:\" на наличие папок source_data и generation_results. """
    directory_list = [r'C:\source_data', r'C:\generation_results']

    for directory in directory_list:
        if os.path.exists(directory) == True:
            print(f'Папка {directory} в директории {os.getcwd()} - создана')
            continue
        else:
            try:
                if directory not in os.getcwd():
                    while True:
                        print(f'Папка {directory} в директори {os.getcwd()} - отсутсвует!')
                        print(f'Создаю папку >>> {directory}\n')
                        os.mkdir(directory)
                        break
                else:
                    return f'Папка {directory} в директори {os.getcwd()} - найдена!'
            except (FileNotFoundError, FileExistsError):
                continue
    return 'Все необходимые папки - найдены'


def check_source_data():
    """Проверка директории "source_data" на наличие исходных файлов. """
    directory = r'C:\source_data'
    files = os.listdir(directory)
    try:
        if len(files) == 0:
            raise EmptyException(f'В папке {directory} нет файлов!')
        else:
            pass
    except EmptyException as error:
        raise SystemExit(error)


def check_generation_results():
    """Проверка директории "generation_results" на наличие итоговых файлов. """
    directory = r'C:\generation_results'
    files = os.listdir(directory)
    try:
        if len(files) == 0:
            pass
        else:
            raise EmptyException('В папке уже есть итоговые файлы!')
    except EmptyException as error:
        raise SystemExit(error)


def check_lists(arg):
    """Проверка листов в книге по образцу и вывод отчета в файл Excel. """

    incorrect_list_of_files_names = []
    incorrect_list_of_files = []
    correct_list_of_files_17 = []
    correct_list_of_files_18 = []
    correct_sheet = [['налог на прибыль', 'НДС', 'НДС импорт ТС'],
                     ['штраф 119 НК РФ']]

    for file in tqdm(arg):
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names

        if sheets == correct_sheet[0]:
            correct_list_of_files_17.append(Path(xl).name)
        elif sheets == correct_sheet[1]:
            correct_list_of_files_18.append(Path(xl).name)
        else:
            incorrect_list_of_files_names.append(Path(xl).name)
            incorrect_list_of_files.append(sheets)
    dict_of_errors = {'Имя файла': incorrect_list_of_files_names,
                      'Наименование листов': incorrect_list_of_files}
    df = pd.DataFrame(dict_of_errors)
    writer = ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl')
    df.to_excel(writer, sheet_name='Отчет по ошибкам')
    writer.close()
    # for broken_file in incorrect_list_of_files:
    #     wb = openpyxl.load_workbook(broken_file)
    #     sheets = wb.sheetnames
    #     sheets.title = 'Налог'
    #     print(sheets[0])
    #     if len(sheets) == 3:
    #         sheets.title = correct_sheet[0]
    #         wb.save(broken_file)
    #     elif len(sheets) == 2:
    #         sheets.title = correct_sheet[0]
    #         wb.save(broken_file)
    #     else:
    #         sheets.title = correct_sheet[1]
    #         wb.save(broken_file)

    print(f'Файлов с ошибкой: {len(incorrect_list_of_files_names)}\n список файлов: {incorrect_list_of_files_names}\n'
          f'Корректных файлов п.17: {len(correct_list_of_files_17)}\n список файлов: {correct_list_of_files_17}\n'
          f'Корректных файлов п.18: {len(correct_list_of_files_18)}\n список файлов: {correct_list_of_files_18}\n')


def head_of_table():
    """Опредление шапки таблиц файлов Excel. """
    pass


def count_empty_rows():
    """Подсчет пустых строк в таблице файлов Excel. """
    pass


def comparison_columns_in_data():
    """Сравнение колонок с шаблоном в таблицах Excel. """
    mistake_dict = {}
    incorrect_dict_of_files = {}
    for file in tqdm(files_source):
        dict_of_inc_list_files = {}
        df = pd.ExcelFile(file)
        sheets = df.sheet_names
        file_name = Path(df).name
        for sheet in range(len(sheets)):
            df_excel_file = pd.read_excel(file,
                                          header=7,
                                          sheet_name=sheet,
                                          dtype='str'
                                          )
            column_to_list = df_excel_file.columns.tolist()
            # словарь второго уровня, который содержит ключ - имя листа, значение - наименование колонок
            incorrect_dict_of_files[SHEET_NAME_122[sheet]] = column_to_list
            # словарь первого уровня, который содержит ключ - имя файла, значение - словарь второго уровня
            dict_of_inc_list_files[file_name] = incorrect_dict_of_files
        # создаем цикл, который перебирает словарь первого уровня,
        # по ключу - имени файла и значению - листы + наименования колонок
        for file_name_key, sheet_name_value in dict_of_inc_list_files.items():
            incorrect_list_of_sheets = []
            for num_of_sheets in range(len(sheet_name_value)):
                if dict_of_inc_list_files[file_name][SHEET_NAME_122[num_of_sheets]] != COL_NAME[num_of_sheets]:
                    incorrect_list_of_sheets.append(SHEET_NAME_122[num_of_sheets])
                    mistake_dict[file_name_key] = incorrect_list_of_sheets
                else:
                    pass
    df = pd.DataFrame.from_dict(mistake_dict,  orient='index')
    df = df.transpose()
    try:
        with ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl',
                         mode='a' if os.path.exists(FILE_NAME_CHECK_REPORT) else 'w') as writer:
            df.to_excel(writer, sheet_name='Отчет о заголовках')
    except Exception:
        pass
    print('Проверка выполнена, см. отчет!')


def comparison_number_of_rows_in_data():
    """Сравнение кол-ва строк в таблицах Excel. """
    incorrect_list_of_files = []
    dict_of_files_name_and_sum = {}
    dict_of_files_name_and_sum_2 = []

    for file in tqdm(files_source):
        df = pd.ExcelFile(file)
        sheet = df.sheet_names
        file_name = str(Path(df).name)
        dict_of_files_name_and_sum[file_name] = {}
        df_excel_file = pd.read_excel(file,
                                      header=7,
                                      sheet_name=sheet[0],
                                      dtype='str'
                                      )
        l = df_excel_file.columns.tolist()
        if l != COL_NAME.tolist():
            incorrect_list_of_files.append(Path(df).name)
        else:
            pass
    return incorrect_list_of_files
    #     dict_of_files_name_and_sum[file_name] = {sheet[0]: file_obj.shape[0]}
    #         for i in range(3):
    #             dict_of_files_name_and_sum[file_name] = {i: file_obj.shape[0]}
    #             # dict_of_files_name_and_sum[file_name][1] = file_obj.shape[0]
    #             # dict_of_files_name_and_sum[file_name][2] = file_obj.shape[0]
    #             # dict_of_files_name_and_sum[file_name] = [sheet[2], file_obj.shape[0]]
    #             dict_of_files_name_and_sum.update()
    # for file in tqdm(files_comp):
    #     df = pd.ExcelFile(file)
    #     sheet = df.sheet_names
    #     file_name = str(Path(df).name)
    #     dict_of_files_name_and_sum_2[file_name] = {}
    #     file_obj = pd.read_excel(file,
    #                              header=7,
    #                              sheet_name=sheet[0],
    #                              dtype='str'
    #                              )
    #     dict_of_files_name_and_sum_2[file_name] = {sheet[0]: file_obj.shape[0]}
    # pprint(dict_of_files_name_and_sum_2)
    # i = 0
    # d = []
    # for key, value in dict_of_files_name_and_sum_2:
    #     for val in dict_of_files_name_and_sum.values():
    #         if value['налог на прибыль'] != val['налог на прибыль']:
    #             i += 1
    #         else:
    #             d.append(key)
    return d


def count_files():
    """Финальный подсчет обрабатываемых файлов в рабочей папке. """
    print(f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx\n"
          f"В папке {main_folder} хранится {len(list(glob.glob('*.xls')))} объектов в формате .xls")
