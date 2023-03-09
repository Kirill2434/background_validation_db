import os.path
from pathlib import Path

import pandas as pd
from pandas import ExcelWriter
from pprint import pprint

from source.config import all_files, FRPV_CHECK_REPORT, FRPV_CHECK_REPORT_N, SHEET_NAME_FRPV, COL_NAME_FRPV, FILE_NAME_CHECK_REPORT
from utils import record_to_excel
from tqdm import tqdm


def check_frpv_sheets(arg):
    """Проверка листов в книге по образцу и вывод отчета в файл Excel. """
    list_of_files = []
    list_of_sheets = []
    correct_files_names = []
    correct_sheet = ['Раздел 1', 'Раздел 2', 'Раздел 3']
    file_number = 1

    for file in tqdm(arg):

        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names
        for cor in correct_sheet:
            if cor not in sheets:
                list_of_files.append(file)
                list_of_sheets.append(cor)
            else:
                 pass
    dict_of_errors = {'Имя файла': list_of_files,
                      'Наименование листов': list_of_sheets}
    df = pd.DataFrame.from_dict(dict_of_errors, orient='index')
    df = df.transpose()
    while True:
        if not os.path.exists(FRPV_CHECK_REPORT):
            writer = ExcelWriter(FRPV_CHECK_REPORT,
                                 engine='openpyxl')
            df.to_excel(writer, sheet_name='Отчет по ошибкам')
            writer.close()
            break
        else:
            new_file_name = FRPV_CHECK_REPORT_N + '_' + str(file_number) + '.xlsx'
            writer = ExcelWriter(new_file_name,
                                 engine='openpyxl')
            df.to_excel(writer, sheet_name='Отчет по ошибкам в листах')
            writer.close()
            file_number += 1
            break
    return None


def comparison_columns_in_frpv(path):
    """Сравнение колонок с шаблоном в таблицах Excel. """
    mistake_dict = {}
    incorrect_dict_of_files = {}
    incorrect_list_of_files = []
    for file in tqdm(path):
        dict_of_inc_list_files = {}
        df = pd.ExcelFile(file)
        sheets = df.sheet_names
        file_name = Path(df).name
        for name in sheets:
            if name == SHEET_NAME_FRPV[0]:
                df_excel_file = pd.read_excel(file,
                                              header=5,
                                              sheet_name=name,
                                              dtype='str'
                                              )
                column_to_list = df_excel_file.columns.tolist()
                incorrect_dict_of_files[SHEET_NAME_FRPV[0]] = column_to_list
                dict_of_inc_list_files[file_name] = incorrect_dict_of_files
            if name == SHEET_NAME_FRPV[1]:
                df_excel_file = pd.read_excel(file,
                                              header=5,
                                              sheet_name=name,
                                              dtype='str'
                                              )
                column_to_list = df_excel_file.columns.tolist()
                incorrect_dict_of_files[SHEET_NAME_FRPV[1]] = column_to_list
                dict_of_inc_list_files[file_name] = incorrect_dict_of_files
            if name == SHEET_NAME_FRPV[2]:
                df_excel_file = pd.read_excel(file,
                                              header=5,
                                              sheet_name=name,
                                              dtype='str'
                                              )
                column_to_list = df_excel_file.columns.tolist()
                incorrect_dict_of_files[SHEET_NAME_FRPV[2]] = column_to_list
                dict_of_inc_list_files[file_name] = incorrect_dict_of_files
        # создаем цикл, который перебирает словарь первого уровня,
        # по ключу - имени файла и значению - листы + наименования колонок
        for file_name_key, sheet_name_value in dict_of_inc_list_files.items():
            incorrect_list_of_sheets = []
            for num_of_sheets in range(len(sheet_name_value)):
                if dict_of_inc_list_files[file_name][SHEET_NAME_FRPV[num_of_sheets]] != COL_NAME_FRPV[num_of_sheets]:
                    incorrect_list_of_sheets.append(SHEET_NAME_FRPV[num_of_sheets])
                    mistake_dict[file_name_key] = incorrect_list_of_sheets
                else:
                    pass
    df_1_page = pd.DataFrame(incorrect_list_of_files)
    df = pd.DataFrame.from_dict(mistake_dict,  orient='index')
    df = df.transpose()
    record_to_excel(df, df_1_page, 'Отчет о заголовках')
    return None


def merge_frpv_files(files):
    """Слияние необходимых листов в файлах Excel в которых содержится множество листов. """
    list_r_1 = []
    list_r_2 = []
    list_r_3 = []
    incorrect_files = []
    correct_sheet = ['Раздел 1', 'Раздел 2', 'Раздел 3']
    for file in tqdm(files):
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names
        for sheet in sheets:
            if sheet == correct_sheet[0]:
                file_obj = pd.read_excel(file,
                                         header=5,
                                         sheet_name=correct_sheet[0],
                                         dtype='str'
                                         )
                file_obj.insert(26, 'Файл источник', file)
                list_r_1.append(file_obj)
            elif sheet == correct_sheet[1]:
                file_obj = pd.read_excel(file,
                                         header=5,
                                         sheet_name=correct_sheet[1],
                                         dtype='str'
                                         )
                file_obj.insert(13, 'Файл источник', file)
                list_r_2.append(file_obj)
            elif sheet == correct_sheet[2]:
                file_obj = pd.read_excel(file,
                                         header=5,
                                         sheet_name=correct_sheet[2],
                                         dtype='str'
                                         )
                file_obj.insert(11, 'Файл источник', file)
                list_r_3.append(file_obj)

    merge_file = pd.concat(list_r_1)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 1.xlsx',
                        sheet_name=correct_sheet[0],
                        index=False)
    merge_file = pd.concat(list_r_2)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 2.xlsx',
                        sheet_name=correct_sheet[1],
                        index=False)
    merge_file = pd.concat(list_r_3)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 3.xlsx',
                        sheet_name=correct_sheet[2],
                        index=False)
    return None

# print(check_frpv_sheets(all_files))
print(merge_frpv_files(all_files))
