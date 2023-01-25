import glob
from datetime import time

import pandas as pd


def pandas_cut(path_file_in, path_dir_out, count_lines=10000):
    """Функция pandas_cut разрезает файл с расширением .xlsx на файлы с расширением .xlsx c заданным количеством строк.

      Параметры: path_file_in : str
                 Абсолютный или относительный путь до файла с расширением .xlsx, который нужно разрезать.
                 path_dir_out : str
                    Абсолютный или относительный путь до папки, в которую будут помещаться нарезанные файлы.
                  count_lines :  int, default 10 000
                    Количество строк, на которые разрезается исходный файл.
       Возвращаемое значение: None
    """

    df = pd.ExcelFile(path_file_in)
    path_out = path_dir_out + '\\' + path_file_in.split('\\')[-1][:-4]
    sheets = df.sheet_names
    file_number = 1
    skiprows = 0
    try:
        while True:
            file_obj = pd.read_excel(path_file_in,
                                     sheet_name=sheets[0],
                                     skiprows=skiprows,
                                     nrows=count_lines,
                                     dtype='str')
            if file_obj.shape[0] == 0:
                break
            else:
                new_file_name = path_out + '_' + str(file_number) + '.xlsx'
                writer = pd.ExcelWriter(new_file_name, engine='openpyxl')
                file_obj.to_excel(writer, index=False)
                writer.close()
                skiprows += count_lines
                file_number += 1
    except:
        pass

