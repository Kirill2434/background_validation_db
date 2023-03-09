"""Файл с исходными переменными. """

import os

import glob
from pathlib import PureWindowsPath

# необходимо создать функцию, которая принимает эталонный файл
# и наполняет исходные переменные данными
# следующим шагом модернизировать методы,
# нужно довести до ума аргументы, которые принимают на себя функции

os.chdir(r'C:\source_data')

pr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
nds = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]
tc = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]
nk = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]
R_1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]
R_2 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
R_3 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
num_of_header = [7, 8]

COL_NAME = [pr, nds, tc]

SHEET_NAME_122 = ['налог на прибыль', 'НДС', 'НДС импорт ТС']
SHEET_NAME_119 = 'штраф 119 НК РФ'
list_of_sheet_names = [SHEET_NAME_122, SHEET_NAME_119]

SHEET_NAME_FRPV = ['Раздел 1', 'Раздел 2', 'Раздел 3']
COL_NAME_FRPV = [R_1, R_2, R_3]

FILE_NAME = r'C:\generation_results\common_data_file.xlsx'
FILE_NAME_99_122 = r'C:\generation_results\common_data_file_99.xlsx'
FILE_NAME_shtraf_119_NK_RF = r'C:\generation_results\common_data_shtraf_119_NK_RF_file.csv'
FILE_NAME_CHECK_REPORT = r'C:\generation_results\check_report_file.xlsx'
FRPV_CHECK_REPORT = r'C:\generation_results\check_report_file.xlsx'
FRPV_CHECK_REPORT_N = r'C:\generation_results\check_report_file'
files_3_pages = glob.glob(r'C:\source_data\???122.xlsx')
files_1_pages = glob.glob(r'C:\source_data\??.xlsx')
files_comp = glob.glob(r'C:\comparision_source_data\???122.xlsx')
check_folder = glob.glob(r'C:\generation_results')
file_paths = [files_3_pages, files_1_pages]

main_folder = PureWindowsPath(r'C:\source_data')


all_files = glob.glob('*.xlsx')
files_nn_122 = glob.glob('???122.xlsx')
files_nn = glob.glob('??.xlsx')
files_nnn = glob.glob('???.xlsx')
files_nnnn_119 = glob.glob('99???119.xlsx')
files_nnnn_122 = glob.glob('99???122.xlsx')
