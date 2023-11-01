import os

file_path = r'C:\new'

for file in os.listdir(file_path):
    for i in os.listdir(fr'{file_path}\{file}'):
        if i.split('.')[1] == 'xlsx':
            os.rename(fr'{file_path}\{file}\{i}', fr"{file_path}\{file}\{i.split('.')[0]}{file}.xlsx")

