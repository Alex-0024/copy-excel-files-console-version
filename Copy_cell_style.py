# Copy_cell_style.py

from copy import copy

# Функция для копирования стилей ячейки
def copy_cell_style(begin, res):
    if begin.has_style:
        res.font = copy(begin.font)
        res.border = copy(begin.border)
        res.fill = copy(begin.fill)
        res.number_format = copy(begin.number_format)
        res.alignment = copy(begin.alignment)