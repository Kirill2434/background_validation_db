import os

from Parsing_excel import excel_merger
from Parsing_excel.config import all_files, files_nn_122, main_folder


class EmptyException(Exception):
    def __init__(self, text):
        self.txt = text


def check_generation_data():
    directory = 'C:\generation_results'
    files = os.listdir(directory)
    try:
        if len(files) == 0:
            pass
        else:
            raise EmptyException('В папке уже есть итоговые файлы!')
    except EmptyException as error:
        raise SystemExit(error)


if __name__ == '__main__':
    check_generation_data()
    print('Старт')
    excel_merger.check_lists(all_files)
    print('Слияние файла: файл №1')
    excel_merger.merge_small_files(files_nn_122)
    print('Слияние файла: файл №2')
    excel_merger.merge_large_files()
    excel_merger.count_files()



