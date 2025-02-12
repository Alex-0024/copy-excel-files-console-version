# Сopy excel files
This application is designed to make it easier to work with Excel files when you need to combine several files into one final file. In my example, the files have the same number of columns and all files have the same first row. After processing, the same rows are colored yellow.
## For example

- Input data

![Image alt](https://github.com/Alex-0024/copy-excel-files-console-version/blob/master/1.png)
- Add all input of files and this application into folder

![Image alt](https://github.com/Alex-0024/copy-excel-files-console-version/blob/master/2.png)
- Working application

![Image alt](https://github.com/Alex-0024/copy-excel-files-console-version/blob/master/3.png)
- Result

![Image alt](https://github.com/Alex-0024/copy-excel-files-console-version/blob/master/4.png)

To build the application, open a new project in PyCharm. Upload files with the extension .py. Upload the necessary libraries. After that, to create an executable file with the extension .exe, install pyinstaller (pip install pyinstaller) and create the necessary file (pyinstaller --onefile main.py)
