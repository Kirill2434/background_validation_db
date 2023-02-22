import os

from pprint import pprint

import pandas as pd
from pathlib import Path

from pandas import ExcelWriter
from tqdm import tqdm

from source.config import (COL_NAME, SHEET_NAME_122, FILE_NAME_CHECK_REPORT,
                           files_3_pages, file_paths, SHEET_NAME_119, nk, files_comp)
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
    incorrect_files = []
    correct_files_1 = []
    correct_files_2 = []
    correct_sheet = ['Раздел 1', 'Раздел 2', 'Раздел 3']

    for file in tqdm(arg):
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names

        if sheets == correct_sheet[0]:
            correct_files_1.append(Path(xl).name)
        elif sheets == correct_sheet[1]:
            correct_files_2.append(Path(xl).name)
        else:
            incorrect_list_of_files_names.append(Path(xl).name)
            incorrect_files.append(sheets)
    dict_of_errors = {'Имя файла': incorrect_list_of_files_names,
                      'Наименование листов': incorrect_files}
    df = pd.DataFrame(dict_of_errors)
    writer = ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl')
    df.to_excel(writer, sheet_name='Отчет по ошибкам')
    writer.close()
    print(f'Файлов с ошибкой: {len(incorrect_list_of_files_names)}\n список файлов: {incorrect_list_of_files_names}\n'
          f'Корректных файлов: {len(correct_files_1)}\n список файлов: {correct_files_1}\n'
          f'Корректных файлов: {len(correct_files_2)}\n список файлов: {correct_files_2}\n')


def comparison_columns_in_data():
    """Сравнение колонок с шаблоном в таблицах Excel. """
    mistake_dict = {}
    incorrect_dict_of_files = {}
    incorrect_list_of_files = []
    for path in file_paths:
        if path == file_paths[1]:
            for file in tqdm(path):
                dict_of_inc_list_files = {}
                df = pd.ExcelFile(file)
                sheets = df.sheet_names
                file_name = Path(df).name
                df_excel_file = pd.read_excel(file,
                                              header=6,
                                              sheet_name=sheets[0],
                                              dtype='str'
                                              )
                # преобразование колонок обрабатываемого датафрейма в список
                column_to_list = df_excel_file.columns.tolist()
                # словарь второго уровня, который содержит ключ - имя листа, значение - наименование колонок
                incorrect_dict_of_files[SHEET_NAME_119] = column_to_list
                # словарь первого уровня, который содержит ключ - имя файла, значение - словарь второго уровня
                dict_of_inc_list_files[file_name] = incorrect_dict_of_files
                # создаем цикл, который перебирает словарь первого уровня,
                # по ключу - имени файла и значению - листы + наименования колонок
                if dict_of_inc_list_files[file_name][SHEET_NAME_119] != nk:
                    incorrect_list_of_files.append(file_name)
                else:
                    pass
        else:
            for file in tqdm(path):
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
                    incorrect_dict_of_files[SHEET_NAME_122[sheet]] = column_to_list
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
    df_1_page = pd.DataFrame(incorrect_list_of_files)
    df = pd.DataFrame.from_dict(mistake_dict,  orient='index')
    df = df.transpose()
    # todo вывести запись в файл в отдельный метод в utils
    # todo если файл уже создан, то просто дописываем на существующий лист данные
    if os.path.exists(FILE_NAME_CHECK_REPORT):
        with ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl',
                         mode='a', if_sheet_exists='overlay'
                         ) as writer:
            df.to_excel(writer, sheet_name='Отчет о заголовках')
            df_1_page.to_excel(writer, sheet_name='Отчет о заголовках', startrow=5)
            print('Проверка выполнена, см. отчет!')

    else:
        with ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl',
                         mode='a' if os.path.exists(FILE_NAME_CHECK_REPORT) else 'w'
                         ) as writer:
            df.to_excel(writer, sheet_name='Отчет о заголовках')
            df_1_page.to_excel(writer, sheet_name='Отчет о заголовках', startrow=5)
            print('Проверка выполнена, см. отчет!')
    return None


def comparison_number_of_rows_in_data():
    """Сравнение кол-ва строк в таблицах Excel. """
    number_of_rows_dict = {}
    number_of_comp_rows_dict = {}
    list_of_result = []
    for file in tqdm(files_3_pages):
        name_of_files_key_dict = {}
        df = pd.ExcelFile(file)
        sheets = df.sheet_names
        file_name = str(Path(df).name)
        for sheet in range(len(sheets)):
            df_excel_file = pd.read_excel(file,
                                          header=7,
                                          sheet_name=sheet,
                                          dtype='str'
                                          )
            df = df_excel_file.shape[0]
            number_of_rows_dict[SHEET_NAME_122[sheet]] = df
            name_of_files_key_dict[file_name] = number_of_rows_dict
        for fi in tqdm(files_comp):
            name_of_comp_files = {}
            df = pd.ExcelFile(fi)
            sheets = df.sheet_names
            file_name = str(Path(df).name)
            for sheet in range(len(sheets)):
                df_excel_file = pd.read_excel(file,
                                              header=7,
                                              sheet_name=sheet,
                                              dtype='str'
                                              )
                df = df_excel_file.shape[0]
                number_of_comp_rows_dict[SHEET_NAME_122[sheet]] = df
                name_of_comp_files[file_name] = number_of_rows_dict
            pprint(name_of_comp_files)
            # if name_of_files_key_dict[file_name][SHEET_NAME_122] != name_of_comp_files[file_name][SHEET_NAME_122]:
            #     list_of_result.append(file_name)
    pprint(list_of_result)
    return None
