import os.path
from pathlib import Path

import pandas as pd
from pandas import ExcelWriter

from source.config import all_files, FRPV_CHECK_REPORT, FRPV_CHECK_REPORT_N
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
            df.to_excel(writer, sheet_name='Отчет по ошибкам')
            writer.close()
            file_number += 1
            break
    return None


def rewrite_sheet_name(files):
    files_list = []
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
                files_list.append(file_obj)
    merge_file = pd.concat(files_list)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 1.xlsx',
                        sheet_name=correct_sheet[0],
                        index=False)
            # elif sheet == correct_sheet[1]:
            #     file_obj = pd.read_excel(file,
            #                              header=5,
            #                              sheet_name=correct_sheet[1],
            #                              dtype='str'
            #                              )
            #     file_obj.insert(13, 'Файл источник', file)
            #     files_list.append(file_obj)
            #     merge_file = pd.concat(files_list)
            #     merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 2.xlsx',
            #                         sheet_name=correct_sheet[1],
            #                         index=False)
            # else:
            #     file_obj = pd.read_excel(file,
            #                              header=5,
            #                              sheet_name=correct_sheet[2],
            #                              dtype='str'
            #                              )
            #     file_obj.insert(11, 'Файл источник', file)
            #     files_list.append(file_obj)
            #     merge_file = pd.concat(files_list)
            #     merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 3.xlsx',
            #                         sheet_name=correct_sheet[2],
            #                         index=False)
    return None



print(rewrite_sheet_name(all_files))
