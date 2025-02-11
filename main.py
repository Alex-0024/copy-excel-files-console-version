# main.py

import sys
from Find_all_excel import find_all_excel
from Copy_files import copy_file, res_path
from Copy_row_column_sizes import copy_row_column_sizes
from Find_the_same_data_and_paint import find_data_and_paint

print('Привет, начинаю работу')

# Нахождение списка файлов Excel в рабочей папке
file_paths = find_all_excel()

print('Высота строк в итоговом файле будет определена по первой строке первого файла в списке')

enter = input('Вы уверены, что хотите объединить эти файлы Yes/No:')
while enter != 'Yes':
    if enter == 'No':
        print('Программа завершена без объединения')
        input()
        sys.exit()
    else:
        enter = input('введите Yes или No:')

# Копирование исходных файлов Excel в общий файл с проверкой формата файлов
try:
    copy_file(file_paths)
except Exception as e:
    print(f'Произошла ошибка, возможно файлы имеют неверный формат или повреждены: {e}')
    print('Программа завершена')
    input()
    sys.exit()

# Установка высоты строк для всего целевого файла на основании первой строки первого файла из списка
# Установка ширины столбцов целевого файла на основании первого файла из списка
copy_row_column_sizes(file_paths[0], res_path)

# Определение строк с одинаковым значением в первой ячейке и их закраска
find_data_and_paint(res_path)

print(f'Работа завершена, файл с именем {res_path} создан успешно')
print('Одинаковые повторяющиеся строки выделены желтым цветом, кроме первой строки')
input()