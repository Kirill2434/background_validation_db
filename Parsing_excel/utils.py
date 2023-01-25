import glob

from source.config import main_folder


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
