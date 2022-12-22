import os

import pandas as pd
import openpyxl as oxp
import glob


FILE_NAME = 'C:\source_data\common_data_file.xlsx'

os.chdir('C:\source_data')
files = glob.glob('*.xlsx')

'''Опредление заголовков таблицы файлов Excel'''

sheet = pd.ExcelFile(files[0])
print(sheet)

