# main.py

import sys
from Find_all_excel import find_all_excel
from Copy_files import copy_file, res_path
from Copy_row_column_sizes import copy_row_column_sizes

print('Привет, начинаю работу')
print('Высота строк в итоговом файле будет определена по первой строке первого файла в списке')

# Нахождение списка файлов Excel в рабочей папке
file_paths = find_all_excel()

enter = input('Вы уверены, что хотите объединить эти файлы Yes/No:')
while enter != 'Yes':
    if enter == 'No':
        print('Программа завершена без объединения')
        input()
        sys.exit()
    else:
        enter = input('введите Yes или No:')

# Копирование исходных файлов Excel в общий файл
copy_file(file_paths)

# Установка высоты строк для всего целевого файла на основании первой строки первого файла из списка
# Установка ширины столбцов целевого файла на основании первого файла из списка
copy_row_column_sizes(file_paths[0], res_path)

print(f'Работа завершена, файл с именем {res_path} создан успешно')
input()