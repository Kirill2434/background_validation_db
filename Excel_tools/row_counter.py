import pandas as pd

path_1 = r'C:\source_data\resurs.xlsx'
path_2 = r'C:\source_data\main.xlsx'

# Файл из которого берем данные
df1 = pd.read_excel(path_1, header=8, sheet_name='Лист1', dtype=str)
# Основной файл в который добовляем данны
df2 = pd.read_excel(path_2,  header=8, sheet_name='Лист1', dtype=str)

unresolved_inn = []
key_num = []

# Записываем данные конкретного стобца в список
list_df1 = df1[3].tolist()
list_df2 = df2[3].tolist()
for inn in list_df1:
    if inn not in list_df2:
        unresolved_inn.append(inn)

match_cell = df1.loc[df1[3].isin(unresolved_inn)]
df2['Статус'] = '-'
df2 = df2.append(match_cell, ignore_index=True)
df2['Статус'] = df2['Статус'].fillna('Новый')
df2.to_excel(r'C:\generation_results\fin.xlsx',
             index=False)


# Работа с несколькими файлами
# files = 'rC:\source_data\*.xlsx'
# main_df = pd.read_excel(path_2,  header=8, sheet_name='Лист1', dtype=str)
# unresolved_inn = []
#
# for file in files:
#     df = pd.read_excel(file,
#                        header=8,
#                        sheet_name='Лист1',
#                        dtype=str)
#     list_of_inn = df[3].tolist()
#     for inn in list_of_inn:
#         if inn not in main_df:
#             unresolved_inn.append(inn)
#
#     match_cell = df.loc[df[3].isin(unresolved_inn)]
#     main_df = main_df.append(match_cell, ignore_index=True)
#     main_df.to_excel(r'C:\generation_results\fin.xlsx',
#                      index=False)
