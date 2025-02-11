# Find_all_excel.py

import os
import sys

def find_all_excel():
    # Получаем список всех файлов в текущей директории
    files = os.listdir('.')

    # Фильтруем файлы с расширением '.xlsx'
    excel_files = [file for file in files if file.endswith('.xlsx')]

    if not excel_files:
        print('Файлов с расширением .xlsx в папке не обнаружено')
        print('Программа завершена')
        input()
        sys.exit()
    else:
        print("Найдены файлы с расширением '.xlsx' в текущей папке:\n")
        for file in excel_files:
            print(file)
        print()

    return excel_files


