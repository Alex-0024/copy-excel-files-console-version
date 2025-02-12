# main.py

import sys
from Find_all_excel import find_all_excel
from Copy_files import copy_file
from Copy_row_column_sizes import copy_row_column_sizes
from Find_the_same_data_and_paint import find_data_and_paint
from Check_count_rows import check_count_rows

res_path = 'result.xlsx'

print('Привет, начинаю работу')

# Нахождение списка файлов Excel в рабочей папке
file_paths = find_all_excel()

print('Высота строк в итоговом файле будет определена по первой строке первого файла в списке')

enter_message = input('Вы уверены, что хотите объединить эти файлы Yes/No:')
while enter_message != 'Yes':
    if enter_message == 'No':
        print('Программа завершена без объединения')
        input()
        sys.exit()
    else:
        enter_message = input('введите Yes или No:')

# Копирование исходных файлов Excel в общий файл с проверкой формата файлов
try:
    print('\tкопирую...')
    copy_file(file_paths, res_path)
except Exception as e:
    print(f'Произошла ОШИБКА !!!, возможно файлы имеют неверный формат или повреждены: {e}')
    print('Программа завершена')
    input()
    sys.exit()

# Установка высоты строк для всего целевого файла на основании первой строки первого файла из списка
# Установка ширины столбцов целевого файла на основании первого файла из списка
print('\tнастраиваю ширину и высоту строк и столбцов...')
copy_row_column_sizes(file_paths[0], res_path)

# Определение строк с одинаковым значением в первой ячейке и их закраска
print('\tнахожу одинаковые первые строки и закрашиваю...')
find_data_and_paint(res_path)

# Проверка соответствия общего кол-ва строк исходных файлов с итоговым файлом
print('\tпровожу проверку правильности копирования...')
if check_count_rows(file_paths, res_path):
    print(f'Работа завершена, итоговый файл {res_path} создан успешно')
    print('Одинаковые повторяющиеся строки выделены желтым цветом, кроме первой строки')
else:
    print(f'Работа завершена, итоговый файл {res_path} имеет ОШИБКИ !!! по количеству скопированных строк')
    print(f'Возможно необходимо удалить файл {res_path} перед копированием')
    print('Повторите операцию')

input()