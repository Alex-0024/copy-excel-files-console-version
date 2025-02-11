# main.py

import sys
from Find_all_excel import find_all_excel
from Copy_files import copy_file, res_path

print('Привет, начинаю работу')
# Список файлов
file_paths = find_all_excel()

enter = input('Вы уверены, что хотите объединить эти файлы Yes/No:')
while enter != 'Yes':
    if enter == 'No':
        print('Программа завершена без объединения')
        input()
        sys.exit()
    else:
        enter = input('введите Yes или No:')

# Копирование файлов
copy_file(file_paths)

print(f'Работа завершена, файл с именем {res_path} создан успешно')
input()