import glob
import os
from pprint import pprint

from pandas import ExcelWriter
import openpyxl as op
from prettytable import PrettyTable

from FRPV.frpv_config import FRPV_CHECK_REPORT_N, FRPV_CHECK_REPORT
from source.config import main_folder


def head_of_table(file):
    """Опредление шапки таблиц файлов Excel. """
    wb = op.load_workbook(file)
    ws = wb['Раздел 1']
    head_of_file = []
    num_of_col = 1
    index_of_head_row = 0
    count_of_none = 0
    for col in ws.iter_rows(min_row=1, max_col=1, max_row=20):
        none_list = []
        for cell in col:
            none_list.append(cell.value)
            for empty_cell in none_list:
                if empty_cell is None:
                    count_of_none += 1
                else:
                    pass
    if count_of_none == 20:
        num_of_col += 1
    else:
        pass
    for row in ws.iter_cols(min_col=num_of_col, max_col=num_of_col):
        for cell in row:
            if cell.value is None:
                pass
            else:
                try:
                    int_cell = int(cell.value)
                    if int_cell == 1:
                        num_of_col += 1
                        for row_second in ws.iter_cols(min_col=num_of_col, max_col=num_of_col):
                            for cell_second in row_second:
                                if cell_second.value is None:
                                    pass
                                else:
                                    try:
                                        int_cell = int(cell_second.value)
                                        if int_cell == 2:
                                            number_of_rows = cell.row
                                            index_of_head_row += number_of_rows
                                        else:
                                            break
                                    except ValueError:
                                        pass
                except ValueError:
                    pass

    for row in ws.iter_rows(min_row=index_of_head_row, max_row=index_of_head_row, values_only=True):
        for cell in row:
            if cell is None:
                pass
            else:
                head_of_file.append(int(cell))
    return head_of_file
    # return 'Fin'


d = {}
all = glob.glob(r'C:\ФРПВ_1 раздел\source_data\*.xlsx')
for a in all:
    d[a] = head_of_table(a)
pprint(d, depth=1)


def count_empty_rows():
    """Подсчет пустых строк в таблице файлов Excel. """
    pass


def count_files():
    """Финальный подсчет обрабатываемых файлов в рабочей папке. """
    xlsx_files = f"В папке {main_folder} хранится {len(list(glob.glob('*.xlsx')))} объектов в формате .xlsx\n"
    xls_files = f"В папке {main_folder} хранится {len(list(glob.glob('*.xls')))} объектов в формате .xls"
    xlsx_and_xls = (xlsx_files, xls_files)
    return xlsx_and_xls



def record_to_excel(obj, sheet_name):
    """Универсальная функция записи результатов выполнения различных
    проверок и действий в других модулях проекта.

    @param obj: датафрейм который необходимо записать в файл формата xlsx
    @param sheet_name: имя листа на который будут записаны данные
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
