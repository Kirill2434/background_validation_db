from Parsing_excel import utils
from source.config import all_files


if __name__ == '__main__':
    # utils.check_directory()
    # utils.check_source_data()
    # utils.check_generation_results()
    # print('Стар блока проверки\n')
    # utils.check_lists(all_files)
    print('Проверка колонок\n')
    utils.comparison_columns_in_data()
    # print('Слияние файла: файл 122\n')
    # excel_merger.merge_small_files()
    # print('Слияние файла: файл 99 122\n')
    # excel_merger.merge_small_files_99()
    # print('Слияние файла: файл №2\n')
    # excel_merger.merge_large_files()
    utils.count_files()
