"""Файл с исходными переменными. """

import os

import glob
from pathlib import PureWindowsPath

os.chdir('C:')

NALOG_NA_PRIBOL = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]
NDS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
NDS_IMPORT_TC = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]
shtraf_119_NK_RF = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]

FILE_NAME = 'C:\generation_results\common_data_file.xlsx'
FILE_NAME_shtraf_119_NK_RF = 'C:\generation_results\common_data_shtraf_119_NK_RF_file.csv'

main_folder = PureWindowsPath('C:\source_data')

all_files = glob.glob('*.xlsx')
files_nn_122 = glob.glob('???122.xlsx')
files_nn = glob.glob('??.xlsx')
files_nnn = glob.glob('???.xlsx')
files_nnnn_nnn = glob.glob('99??????.xlsx')
