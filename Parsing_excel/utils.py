import glob
import os

from pandas import ExcelWriter

from source.config import main_folder, FILE_NAME_CHECK_REPORT


def head_of_table():
    """Опредление шапки таблиц файлов Excel. """
    pass


def count_empty_rows():
    """Подсчет пустых строк в таблице файлов Excel. """
    pass


def count_files():
    """Финальный подсчет обрабатываемых файлов в рабочей папке. """
    print(f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx\n"
          f"В папке {main_folder} хранится {len(list(glob.glob('*.xls')))} объектов в формате .xls")


def record_to_excel(obj_1, obj_2, sheet_name):
    if os.path.exists(FILE_NAME_CHECK_REPORT):
        with ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl',
                         mode='a', if_sheet_exists='overlay'
                         ) as writer:
            obj_1.to_excel(writer, sheet_name=sheet_name)
            obj_2.to_excel(writer, sheet_name=sheet_name, startrow=5)
            print('Проверка выполнена, см. отчет!')

    else:
        with ExcelWriter(FILE_NAME_CHECK_REPORT,
                         engine='openpyxl',
                         mode='a' if os.path.exists(FILE_NAME_CHECK_REPORT) else 'w'
                         ) as writer:
            obj_1.to_excel(writer, sheet_name=sheet_name)
            obj_2.to_excel(writer, sheet_name=sheet_name, startrow=5)
            print('Проверка выполнена, см. отчет!')