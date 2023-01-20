import glob
import os
from pathlib import Path
from pprint import pprint

import pandas as pd
from pandas import ExcelWriter

from tqdm import tqdm

from source.config import (files_nn, files_nnn, FILE_NAME, FILE_NAME_shtraf_119_NK_RF,
                           files_nnnn_119, files_nn_122, files_nnnn_122, FILE_NAME_99_122)


def merge_small_files():
    """Процесс формиования DataFrame'ов путем слияния набора Excel файлов с 3-мя листами"""
    merge_files_nalog_list = []
    merge_files_nds_list = []
    merge_files_nds_import_tc_list = []
    try:
        # Слияние всех листов №1
        for file in tqdm(files_nn_122, desc='Начало слияния листа (налог на прибыль)'):
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
        for file in tqdm(files_nn_122, desc='Начало слияния листа (НДС)'):
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
        for file in tqdm(files_nn_122, desc='Начало слияния листа (НДС импорт ТС)'):
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
        MERGE_FILES = {'Налог на прибыль': merge_files_nalog_na_pribol_merge,
                       'НДС': merge_files_nds_merge,
                       'НДС импорт ТС': merge_files_nds_import_tc_merge}

        writer = pd.ExcelWriter(FILE_NAME, engine='openpyxl')
        for sheet_name in MERGE_FILES.keys():
            MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        return writer.close()
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


def merge_small_files_99():
    """Процесс формиования DataFrame'ов путем слияния набора Excel файлов с 2-мя листами"""
    merge_files_nalog_list = []
    merge_files_nds_list = []
    try:
        # Слияние всех листов №1
        for file in tqdm(files_nnnn_122, desc='Начало слияния листа (налог на прибыль)'):
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
        for file in tqdm(files_nnnn_122, desc='Начало слияния листа (НДС)'):
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
        MERGE_FILES = {'Налог на прибыль': merge_files_nalog_na_pribol_merge,
                       'НДС': merge_files_nds_merge}

        writer = pd.ExcelWriter(FILE_NAME_99_122, engine='openpyxl')
        for sheet_name in MERGE_FILES.keys():
            MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        return writer.close()
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


def merge_large_files():
    """Функция слияния крупных файлов в фомат csv. """
    files_list = [files_nn, files_nnn, files_nnnn_119]
    merge_files_shtraf_119_NK_RF_list = []

    try:
        for dirc in files_list:
            for file in tqdm(dirc, desc='Начало слияния листа (штраф 119 НК РФ)'):
                file_obj_1 = pd.read_excel(file,
                                           header=6,
                                           dtype='str'
                                           )
                file_obj_1[18].replace('\r\n', '\\n', regex=True)
                file_obj_1.insert(18, 'Файл источник', file)
                merge_files_shtraf_119_NK_RF_list.append(file_obj_1)

        merge_files_shtraf_119_NK_RF = pd.concat(merge_files_shtraf_119_NK_RF_list)
        # writer = pd.ExcelWriter(FILE_NAME_shtraf_119_NK_RF, engine='openpyxl')
        # merge_files_shtraf_119_NK_RF.to_excel(writer, sheet_name='штраф 119 НК РФ', index=False)
        # return writer.close()
        return merge_files_shtraf_119_NK_RF.to_csv(FILE_NAME_shtraf_119_NK_RF,
                                                   sep=',',
                                                   index=False,
                                                   encoding='cp1251',
                                                   date_format='str'
                                                   )
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')
