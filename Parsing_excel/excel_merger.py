import os
from pathlib import PureWindowsPath, Path

import pandas as pd
import glob

from tqdm import tqdm

NALOG_NA_PRIBOL = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]
NDS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
NDS_IMPORT_TC = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
shtraf_119_NK_RF = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]

FILE_NAME = 'C:\generation_results\common_data_file.xlsx'
FILE_NAME_shtraf_119_NK_RF = 'C:\generation_results\common_data_shtraf_119_NK_RF_file.csv'

main_folder = PureWindowsPath('C:\source_data')
main_dir = os.chdir('C:\source_data')

all_files = glob.glob('*.xlsx')
files_nn_122 = glob.glob('???122.xlsx')
files_nn = glob.glob('??.xlsx')
files_nnn = glob.glob('???.xlsx')
files_nnnn_nnn = glob.glob('99??????.xlsx')

# Подсчет обрабатываемых файлов в рабочей папке
# print(f"В папке {main_folder} хранится {len(list(all_files))} объектов в формате .xlsx")


files_list = [files_nn, files_nnn, files_nnnn_nnn]

# Проверка листов в книге по образцу
incoorect_list_of_files = []
correct_list_of_files_17 = []
correct_list_of_files_18 = []
for file in tqdm(all_files):
    correct_sheet = ['налог на прибыль', 'НДС', 'НДС импорт ТС']
    correct_sheet_2 = ['штраф 119 НК РФ']
    xl = pd.ExcelFile(file)
    sheets = xl.sheet_names

    if sheets == correct_sheet:
        correct_list_of_files_17.append(Path(xl).name)
    elif sheets == correct_sheet_2:
        correct_list_of_files_18.append(Path(xl).name)
    else:
        incoorect_list_of_files.append(Path(xl).name)

print(f'Файлов с ошибкой: {len(incoorect_list_of_files)}\n список файлов: {incoorect_list_of_files}')
print(f'Корректных файлов п.17: {len(correct_list_of_files_17)}\n список файлов: {correct_list_of_files_17}')
print(f'Корректных файлов п.17: {len(correct_list_of_files_18)}\n список файлов: {correct_list_of_files_18}')

if files_nn_122:
    '''Процесс формиования DataFrame'ов путем слияния набора Excel файлов'''

    # Слияние всех листов №1
    merge_files_nalog_list = []
    for file in tqdm(files_nn_122, desc='Начало слияния листа (налог на прибыль)'):
        sheet = pd.ExcelFile(files_nn_122[0])
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
    merge_files_nds_list = []
    for file in tqdm(files_nn_122, desc='Начало слияния листа (НДС)'):
        sheet = pd.ExcelFile(files_nn_122[0])
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
    merge_files_nds_import_tc_list = []
    for file in tqdm(files_nn_122, desc='Начало слияния листа (НДС импорт ТС)'):
        sheet = pd.ExcelFile(files_nn_122[0])
        sheet = sheet.sheet_names
        file_obj = pd.read_excel(file,
                                 skiprows=7,
                                 usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                                 sheet_name=sheet[2])
        file_obj.columns = [NDS_IMPORT_TC]
        file_obj['Файл источник'] = file
        merge_files_nds_import_tc_list.append(file_obj)
    merge_files_nds_import_tc_merge = pd.concat(merge_files_nds_import_tc_list)

    '''Распредление DataFrame'ов по листам Excel'''

    MERGE_FILES = {'Налог на прибыль': merge_files_nalog_na_pribol_merge,
                   'НДС': merge_files_nds_merge,
                   'НДС импорт ТС': merge_files_nds_import_tc_merge}

    writer = pd.ExcelWriter(FILE_NAME, engine='openpyxl')
    try:
        for sheet_name in MERGE_FILES.keys():
            MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name)

        writer.save()
    except Exception as error:
        print(f'Ошибка при выполнении слияния: {error}')


def merge_large_files():
    """Функция слияния крупных файлов в фомат csv"""
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
        return merge_files_shtraf_119_NK_RF_list
    except Exception as error:
        print(f'Ну мы пытались, но чет пошло не так... {error}')


merge_files_shtraf_119_NK_RF = pd.concat(merge_large_files())
print(merge_files_shtraf_119_NK_RF.to_csv(FILE_NAME_shtraf_119_NK_RF, sep=',', encoding='cp1251'))

# Финальный подсчет обрабатываемых файлов в рабочей папке
print(f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx")
