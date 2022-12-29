import os
import glob
from pathlib import Path

import pandas as pd

from tqdm import tqdm

from Parsing_excel.config import (files_nn, files_nnn, files_nnnn_nnn, NALOG_NA_PRIBOL, NDS,
                                  NDS_IMPORT_TC, FILE_NAME, shtraf_119_NK_RF, FILE_NAME_shtraf_119_NK_RF, main_folder)


def check_lists(arg):
    """Проверка листов в книге по образцу"""

    incoorect_list_of_files = []
    correct_list_of_files_17 = []
    correct_list_of_files_18 = []
    correct_sheet = ['налог на прибыль', 'НДС', 'НДС импорт ТС']
    correct_sheet_2 = ['штраф 119 НК РФ']

    for file in tqdm(arg):
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names

        if sheets == correct_sheet:
            correct_list_of_files_17.append(Path(xl).name)
        elif sheets == correct_sheet_2:
            correct_list_of_files_18.append(Path(xl).name)
        else:
            incoorect_list_of_files.append(Path(xl).name)

        for col in incoorect_list_of_files:
            df = pd.DataFrame(col)
        df.to_excel('output.xlsx')
        return df

    # return (f'Файлов с ошибкой: {len(incoorect_list_of_files)}\n список файлов: {incoorect_list_of_files}\n'
    #         f'Корректных файлов п.17: {len(correct_list_of_files_17)}\n список файлов: {correct_list_of_files_17}\n'
    #         f'Корректных файлов п.17: {len(correct_list_of_files_18)}\n список файлов: {correct_list_of_files_18}\n')


def merge_small_files(arg):
    """Процесс формиования DataFrame'ов путем слияния набора Excel файлов"""
    merge_files_nalog_list = []
    merge_files_nds_list = []
    merge_files_nds_import_tc_list = []
    try:
        # Слияние всех листов №1
        for file in tqdm(arg, desc='Начало слияния листа (налог на прибыль)'):
            sheet = pd.ExcelFile(arg[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     skiprows=7,
                                     usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22],
                                     sheet_name=sheet[0])
            file_obj.columns = [NALOG_NA_PRIBOL]
            file_obj['Файл источник'] = file
            merge_files_nalog_list.append(file_obj)
        merge_files_nalog_na_pribol_merge = pd.concat(merge_files_nalog_list)

        # Слияние всех листов №2
        for file in tqdm(arg, desc='Начало слияния листа (НДС)'):
            sheet = pd.ExcelFile(arg[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     skiprows=7,
                                     usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                                     sheet_name=sheet[1])
            file_obj.columns = [NDS]
            file_obj['Файл источник'] = file
            merge_files_nds_list.append(file_obj)
        merge_files_nds_merge = pd.concat(merge_files_nds_list)

        # Слияние всех листов №3
        for file in tqdm(arg, desc='Начало слияния листа (НДС импорт ТС)'):
            sheet = pd.ExcelFile(arg[0])
            sheet = sheet.sheet_names
            file_obj = pd.read_excel(file,
                                     skiprows=7,
                                     usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                                     sheet_name=sheet[2])
            file_obj.columns = [NDS_IMPORT_TC]
            file_obj['Файл источник'] = file
            merge_files_nds_import_tc_list.append(file_obj)
        merge_files_nds_import_tc_merge = pd.concat(merge_files_nds_import_tc_list)

        # Распредление DataFrame'ов по листам Excel
        MERGE_FILES = {'Налог на прибыль': merge_files_nalog_na_pribol_merge,
                       'НДС': merge_files_nds_merge,
                       'НДС импорт ТС': merge_files_nds_import_tc_merge}

        writer = pd.ExcelWriter(FILE_NAME, engine='openpyxl')
        for sheet_name in MERGE_FILES.keys():
            MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name)
        return writer.close()
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


def merge_large_files():
    """Функция слияния крупных файлов в фомат csv. """
    files_list = [files_nn, files_nnn, files_nnnn_nnn]
    merge_files_shtraf_119_NK_RF_list = []
    try:
        for dir_file in files_list:
            for file in tqdm(dir_file, desc='Начало слияния листа (штраф 119 НК РФ)'):
                file_obj_1 = pd.read_excel(file,
                                           skiprows=7,
                                           usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
                                           )
                file_obj_1.columns = [shtraf_119_NK_RF]
                file_obj_1['Файл источник'] = file
                merge_files_shtraf_119_NK_RF_list.append(file_obj_1)
        merge_files_shtraf_119_NK_RF = pd.concat(merge_files_shtraf_119_NK_RF_list)
        return merge_files_shtraf_119_NK_RF.to_csv(FILE_NAME_shtraf_119_NK_RF, sep=',', encoding='cp1251')
    except Exception as error:
        print(f'Ну мы пытались, но что-то пошло не так... {error}')


def count_files():
    """Финальный подсчет обрабатываемых файлов в рабочей папке. """
    return f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx"
