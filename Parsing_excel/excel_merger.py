import os

import pandas as pd
import glob


NALOG_NA_PRIBOL = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]
NDS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
NDS_IMPORT_TC = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
FILE_NAME = 'C:\source_data\common_data_file.xlsx'

os.chdir('C:\source_data')
files = glob.glob('*.xlsx')

'''Процесс формиования DataFrame'ов путем слияния набора Excel файлов'''

# Слияние всех листов №1
merge_files_nalog_na_pribol = pd.DataFrame()

for file in files:
    sheet = pd.ExcelFile(files[0])
    sheet = sheet.sheet_names
    file_obj = pd.read_excel(file,
                             skiprows=7,
                             usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22],
                             sheet_name=sheet[0])
    file_obj.columns = [NALOG_NA_PRIBOL]
    merge_files_nalog_na_pribol = pd.concat([merge_files_nalog_na_pribol, file_obj])

# Слияние всех листов №1
merge_files_nds = pd.DataFrame()

for file in files:
    sheet = pd.ExcelFile(files[0])
    sheet = sheet.sheet_names
    file_obj = pd.read_excel(file,
                             skiprows=7,
                             usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                             sheet_name=sheet[1])
    file_obj.columns = [NDS]
    merge_files_nds = pd.concat([merge_files_nds, file_obj])

# Слияние всех листов №1
merge_files_nds_import_tc = pd.DataFrame()

for file in files:
    sheet = pd.ExcelFile(files[0])
    sheet = sheet.sheet_names
    file_obj = pd.read_excel(file,
                             skiprows=7,
                             usecols=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20],
                             sheet_name=sheet[2])
    file_obj.columns = [NDS_IMPORT_TC]
    merge_files_nds_import_tc = pd.concat([merge_files_nds_import_tc, file_obj])

'''Распредление DataFrame'ов по листам Excel'''

MERGE_FILES = {'Налог на прибыль': merge_files_nalog_na_pribol,
               'НДС': merge_files_nds,
               'НДС импорт ТС': merge_files_nds_import_tc}

writer = pd.ExcelWriter(FILE_NAME, engine='openpyxl')
try:
    for sheet_name in MERGE_FILES.keys():
        MERGE_FILES[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)

    writer.save()
    print('Слияние файлов успешно выполненно!')
except Exception as error:
    print(f'Ошибка при выполнении слияния: {error}')
