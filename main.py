import os

from Parsing_excel import excel_merger
from source.config import all_files, files_nn_122
from source.custom_exceptions import EmptyException


def check_directory():
    """Проверка директории "С:\" на наличие папок source_data и generation_results. """
    directory_list = ['C:\source_data', 'C:\generation_results']

    for directory in directory_list:
        if os.path.exists(directory) == True:
            print(f'Папка {directory} в директории {os.getcwd()} - создана')
            continue
        else:
            try:
                if directory not in os.getcwd():
                    while True:
                        print(f'Папка {directory} в директори {os.getcwd()} - отсутсвует!')
                        print(f'Создаю папку >>> {directory}\n')
                        os.mkdir(directory)
                        break
                else:
                    return f'Папка {directory} в директори {os.getcwd()} - найдена!'
            except (FileNotFoundError, FileExistsError):
                continue
    return 'Все необходимые папки - найдены'


def check_source_data():
    """Проверка директории "source_data" на наличие исходных файлов. """
    directory = 'C:\source_data'
    files = os.listdir(directory)
    try:
        if len(files) == 0:
            raise EmptyException(f'В папке {directory} нет файлов!')
        else:
            pass
    except EmptyException as error:
        raise SystemExit(error)


def check_generation_results():
    """Проверка директории "generation_results" на наличие итоговых файлов. """
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
    check_directory()
    check_source_data()
    check_generation_results()
    print('Старт')
    # excel_merger.check_lists(all_files)
    print('Слияние файла: файл №1')
    excel_merger.merge_small_files(files_nn_122)
    # print('Слияние файла: файл №2')
    # excel_merger.merge_large_files()
    excel_merger.count_files()
