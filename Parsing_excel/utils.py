import glob
import os

from pandas import ExcelWriter

from FRPV.frpv_config import FRPV_CHECK_REPORT_N, FRPV_CHECK_REPORT
from source.config import main_folder


def head_of_table():
    """Опредление шапки таблиц файлов Excel. """
    pass


def count_empty_rows():
    """Подсчет пустых строк в таблице файлов Excel. """
    pass


def count_files():
    # """Финальный подсчет обрабатываемых файлов в рабочей папке. """
    print(f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx\n"
          f"В папке {main_folder} хранится {len(list(glob.glob('*.xls')))} объектов в формате .xls")


def record_to_excel(obj, sheet_name):
    """Универсальная функция записи результатов выполнения различных
    проверок и дейсвий других модулях проекта.

    @param obj: датафрейм который необходимо записать в файл формата xlsx
    @param sheet_name: имя листа на котрый будут аписаны данные
    """

    # если файл-отчет отсувует в папке, то записываем первый файл-отчет
    if os.path.exists(FRPV_CHECK_REPORT) is False:
        with ExcelWriter(FRPV_CHECK_REPORT,
                         engine='openpyxl') as writer:
            obj.to_excel(writer, sheet_name=sheet_name)
    # когда файл уже есть:
    else:
        # счиаем кол-во файлов-отчетов в папке
        sum_of_files = len(list(glob.glob(r'C:\generation_results\check_report_file??.xlsx')))
        num = sum_of_files + 1
        # создаем новое имя файла с учетом нового номера
        new_file_name = FRPV_CHECK_REPORT_N + '_' + str(num) + '.xlsx'
        with ExcelWriter(new_file_name,
                         engine='openpyxl') as writer:
            obj.to_excel(writer, sheet_name=sheet_name)
