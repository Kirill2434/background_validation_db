import glob
import os

from pandas import ExcelWriter

from FRPV.frpv_config import FRPV_CHECK_REPORT_N
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


def record_to_excel(obj, sheet_name):
    file_number = 0
    while True:
        if os.path.exists(FRPV_CHECK_REPORT_N):
            with ExcelWriter(FILE_NAME_CHECK_REPORT,
                             engine='openpyxl',
                             mode='a', if_sheet_exists='overlay'
                             ) as writer:
                obj.to_excel(writer, sheet_name=sheet_name)
        else:
            file_number += 1
            new_file_name = FRPV_CHECK_REPORT_N + '_' + str(file_number) + '.xlsx'
            with ExcelWriter(new_file_name,
                             engine='openpyxl',
                             mode='a' if os.path.exists(FILE_NAME_CHECK_REPORT) else 'w'
                             ) as writer:
                obj.to_excel(writer, sheet_name=sheet_name)
            break

# ----------------------------
#         file_number = 0
#         while True:
#             if not os.path.exists(FRPV_CHECK_REPORT_N):
#                 writer = ExcelWriter(FRPV_CHECK_REPORT,
#                                      engine='openpyxl')
#                 df.to_excel(writer, sheet_name='Отчет по ошибкам')
#                 writer.close()
#             else:
#                 file_number += 1
#                 new_file_name = FRPV_CHECK_REPORT_N + '_' + str(file_number) + '.xlsx'
#                 writer = ExcelWriter(new_file_name,
#                                      engine='openpyxl')
#                 df.to_excel(writer, sheet_name='Отчет по ошибкам в листах')
#                 writer.close()
#             break
