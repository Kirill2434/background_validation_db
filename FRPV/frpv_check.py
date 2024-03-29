import pandas as pd

from pathlib import Path
from tqdm import tqdm
from frpv_config import SHEET_NAME_FRPV, COL_NAME_FRPV, all_frpv_files


def check_frpv_sheets(path, sheet_name):
    """Проверка листов в книге по образцу и вывод отчета в файл Excel. """
    list_of_files = []
    list_of_sheets = []
    correct_sheet = sheet_name
    for file in tqdm(path):
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
    df.to_excel(r'C:\generation_results\check_report_file.xlsx', 'Отчет по ошибкам в листах', engine='openpyxl')
    return 'Провекра успешно завершена.'


def check_columns_in_frpv(path):
    """Проверка колонок с шаблоном в таблицах Excel. """
    mistake_dict = {}
    incorrect_dict_of_files = {}
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
    df = pd.DataFrame.from_dict(mistake_dict, orient='index')
    df = df.transpose()
    # записываем incorrect_list_of_files в excel отчет по ошибкам
    df.to_excel(r'C:\generation_results\check_col_report_file.xlsx', 'Отчет по ошибкам в шапке', engine='openpyxl')
    return 'Провекра успешно завершена.'

# строка с check_frpv_sheets() проверяет листы xlsx файлов в папке source_data.
# шаг 2, проверить листы в файлах
# print(check_frpv_sheets(all_frpv_files, SHEET_NAME_FRPV))

# строка с check_columns_in_frpv() проверяет числовую шапку в xlsx файлах в папке source_data.
# Кол-во колонок можно отредактировать в файле frpv_config. Переменные R_1, R_2, R_3
# шаг 3, проверить шапку в файлах
print(check_columns_in_frpv(all_frpv_files))
