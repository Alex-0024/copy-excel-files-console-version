import os

def find_all_excel():
    # Получаем список всех файлов в текущей директории
    files = os.listdir('.')

    # Фильтруем файлы с расширением '.xlsx'
    excel_files = [file for file in files if file.endswith('.xlsx')]

    print("Найдены файлы с расширением '.xlsx' в текущей папке:\n")
    for file in excel_files:
        print(file)
    print()

    return excel_files


