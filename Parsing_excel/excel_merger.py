from datetime import date

import pandas as pd
from pyexcelerate import Workbook

from tqdm import tqdm

from source.config import (files_nn, files_nnn, FILE_NAME, FILE_NAME_XLSX,
                           files_nnnn_119, files_nn_122, files_nnnn_122, FILE_NAME_99_122, all_files, files_1_pages)


def merge_small_files():
    """Процесс формиования DataFrame'ов путем слияния набора Excel файлов с 3-мя листами"""
    merge_files_nalog_list = []
    merge_files_nds_list = []
    merge_files_nds_import_tc_list = []
    try:
        # Слияние всех листов №1
        for file in tqdm(all_files, desc='Начало слияния 1 листа'):
            sheet = pd.ExcelFile(files_nn_122[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     header=7,
                                     sheet_name=sheet[0],
                                     dtype='str'
                                     )
            file_obj.insert(23, 'Файл источник', file)
            merge_files_nalog_list.append(file_obj)
        merge_files_nalog_na_pribol_merge = pd.concat(merge_files_nalog_list, ignore_index=False)

        # Слияние всех листов №2
        for file in tqdm(files_nn_122, desc='Начало слияния 2 листа'):
            sheet = pd.ExcelFile(files_nn_122[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     header=7,
                                     sheet_name=sheet[1],
                                     dtype='str'
                                     )
            file_obj.insert(21, 'Файл источник', file)
            merge_files_nds_list.append(file_obj)
        merge_files_nds_merge = pd.concat(merge_files_nds_list)

        # Слияние всех листов №3
        for file in tqdm(files_nn_122, desc='Начало слияния 3 листа'):
            sheet = pd.ExcelFile(files_nn_122[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     header=7,
                                     sheet_name=sheet[2],
                                     dtype='str'
                                     )
            file_obj.insert(21, 'Файл источник', file)
            merge_files_nds_import_tc_list.append(file_obj)
        merge_files_nds_import_tc_merge = pd.concat(merge_files_nds_import_tc_list)

        # Распредление DataFrame'ов по листам Excel
        MERGE_FILES = {'Лист 1': merge_files_nalog_na_pribol_merge,
                       'Лист 2': merge_files_nds_merge,
                       'Лист 3': merge_files_nds_import_tc_merge}

        writer = pd.ExcelWriter(FILE_NAME, engine='openpyxl')
        for sheet_name in MERGE_FILES.keys():
            MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        return writer.close()
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


def merge_small_files_2():
    """Процесс формиования DataFrame'ов путем слияния набора Excel файлов с 2-мя листами"""
    merge_files_nalog_list = []
    merge_files_nds_list = []
    try:
        # Слияние всех листов №1
        for file in tqdm(files_nnnn_122, desc='Начало слияния 1 листа'):
            sheet = pd.ExcelFile(files_nnnn_122[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     header=7,
                                     sheet_name=sheet[0],
                                     dtype='str'
                                     )
            file_obj.insert(23, 'Файл источник', file)
            merge_files_nalog_list.append(file_obj)
        merge_files_nalog_na_pribol_merge = pd.concat(merge_files_nalog_list, ignore_index=False)

        # Слияние всех листов №2
        for file in tqdm(files_nnnn_122, desc='Начало слияния 2 листа'):
            sheet = pd.ExcelFile(files_nnnn_122[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     header=7,
                                     sheet_name=sheet[1],
                                     dtype='str'
                                     )
            file_obj.insert(21, 'Файл источник', file)
            merge_files_nds_list.append(file_obj)
        merge_files_nds_merge = pd.concat(merge_files_nds_list)

        # Распредление DataFrame'ов по листам Excel
        MERGE_FILES = {'Лист 1': merge_files_nalog_na_pribol_merge,
                       'Лист 2': merge_files_nds_merge}

        writer = pd.ExcelWriter(FILE_NAME_99_122, engine='openpyxl')
        for sheet_name in MERGE_FILES.keys():
            MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        return writer.close()
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


def merge_large_files():
    """Функция слияния больших файлов. """
    merge_files = []
    try:
        for file in tqdm(files_1_pages, desc='Начало слияния'):
            file_obj = pd.read_excel(file,
                                     sheet_name='1',
                                     dtype='str')
            file_obj.insert(44, 'Файл источник', file)
            merge_files.append(file_obj)

        merge_files_concat = pd.concat(merge_files)
        # merge_files_concat.to_excel(FILE_NAME_XLSX, sheet_name='Результат',
        #                             index=False)
        # return 'Слияние выполнено'
        # writer = pd.ExcelWriter(FILE_NAME_XLSX, engine='openpyxl')
        # merge_files_concat.to_excel(writer, sheet_name='Результат', index=False)
        # return writer.close()
        return merge_files_concat.to_csv(FILE_NAME_XLSX,
                                         sep='|',
                                         index=False,
                                         encoding='cp1251',
                                         date_format='str'
                                         )
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


# print(merge_large_files())
