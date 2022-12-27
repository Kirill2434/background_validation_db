import os

import pandas as pd
import glob

from tqdm import tqdm

NALOG_NA_PRIBOL = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]
NDS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
NDS_IMPORT_TC = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
shtraf_119_NK_RF = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]

FILE_NAME = 'C:\source_data\common_data_file_2.xlsx'
FILE_NAME_shtraf_119_NK_RF = 'C:\source_data\common_data_shtraf_119_NK_RF_file_2.csv'

main_dir = os.chdir('C:\source_data')

files_nn_122 = glob.glob('???122.xlsx')
files_nn = glob.glob('??.xlsx')
files_nnn = glob.glob('???.xlsx')
files_nnnn_nnn = glob.glob('99??????.xlsx')

files_list = [files_nn, files_nnn, files_nnnn_nnn]

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

# if files_nn or files_nnn or files_nnnn_nnn:
    # Слияние всех листов для файлов с штрафами


def merge_files(dir_file):
    merge_files_shtraf_119_NK_RF_list = []
    try:
        for file in tqdm(dir_file, desc='Начало слияния листа (штраф 119 НК РФ)'):
            file_obj_1 = pd.read_excel(file,
                                       skiprows=7,
                                       usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
                                       )
            file_obj_1.columns = [shtraf_119_NK_RF]
            file_obj_1['Файл источник'] = file
            merge_files_shtraf_119_NK_RF_list.append(file_obj_1)
            # for file_2 in tqdm(files_nnnn_nnn, desc='Начало слияния листа (штраф 119 НК РФ nnnn_nnn)'):
            #     file_obj_2 = pd.read_excel(file_2,
            #                                skiprows=7,
            #                                usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
            #                                )
            #     file_obj_2.columns = [shtraf_119_NK_RF]
            #     file_obj_2['Файл источник'] = file_2
            #     merge_files_shtraf_119_NK_RF_list.append(file_obj_2)
        merge_files_shtraf_119_NK_RF_merge = pd.concat(merge_files_shtraf_119_NK_RF_list)
        tqdm.write('Слияние файлов прошло успешно!')
        return merge_files_shtraf_119_NK_RF_merge
    except Exception as error:
        print(f'Ну мы пытались, но чет пошло не так... {error}')


for dir in files_list:
    print(merge_files(dir).to_csv(FILE_NAME_shtraf_119_NK_RF, sep=',', encoding='cp1251'))
