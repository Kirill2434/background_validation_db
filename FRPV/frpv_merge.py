import os.path
from pathlib import Path

import glob
import pandas as pd
from pandas import ExcelWriter
from pprint import pprint

from frpv_config import all_frpv_files, main_folder
from Parsing_excel.utils import record_to_excel
from tqdm import tqdm


def merge_frpv_files(files):
    """Слияние необходимых листов в файлах Excel в которых содержится множество листов. """
    list_r_1 = []
    list_r_2 = []
    list_r_3 = []
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
                list_r_1.append(file_obj)
            elif sheet == correct_sheet[1]:
                file_obj = pd.read_excel(file,
                                         header=5,
                                         sheet_name=correct_sheet[1],
                                         dtype='str'
                                         )
                file_obj.insert(17, 'Файл источник', file)
                list_r_2.append(file_obj)
            elif sheet == correct_sheet[2]:
                file_obj = pd.read_excel(file,
                                         header=5,
                                         sheet_name=correct_sheet[2],
                                         dtype='str'
                                         )
                file_obj.insert(17, 'Файл источник', file)
                list_r_3.append(file_obj)

    merge_file = pd.concat(list_r_1)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 1.xlsx',
                        sheet_name=correct_sheet[0],
                        index=False)
    merge_file = pd.concat(list_r_2)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 2.xlsx',
                        sheet_name=correct_sheet[1],
                        index=False)
    merge_file = pd.concat(list_r_3)
    merge_file.to_excel(r'C:\generation_results\ФРПВ раздел 3.xlsx',
                        sheet_name=correct_sheet[2],
                        index=False)
    return None


def count_files():
    """Финальный подсчет обрабатываемых файлов в рабочей папке. """
    print(f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx\n"
          f"В папке {main_folder} хранится {len(list(glob.glob('*.xls')))} объектов в формате .xls\n"
          f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsb')))} объектов в формате .xlsb")

# что-бы запустить скрипт, нужно раскомментировать строку ниже, где прописано ключевое слово print
# шаг 1, проверить папку с файлами
# print(count_files())
# строка с функцией merge_frpv_files() запускает слияние подготовленных фрпв файлов
# шаг 4, сляине файлов
# print(merge_frpv_files(all_frpv_files))
