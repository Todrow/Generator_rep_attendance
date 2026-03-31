from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
import datetime
import os

# Сопоставление дней недели и файлов
days_to_files = {
    'Понедельник': None,  # Нет файла
    'Вторник': 'Вторник.xlsx',
    'Среда': 'Среда.xlsx',
    'Четверг': 'Четверг.xlsx',
    'Пятница': 'Пятница.xlsx',      # Нет файла
    'Суббота': 'Суббота.xlsx',
    'Воскресенье': None   # Нет файла
}

# Получаем текущий день недели по-русски
days_russian = ['Понедельник', 'Вторник', 'Среда',
                'Четверг', 'Пятница', 'Суббота', 'Воскресенье']
today = datetime.datetime.now().date()
day_of_week = days_russian[today.weekday()]
filename = days_to_files[day_of_week]

if filename:
    # Загружаем сохранённый файл через openpyxl
    wb = load_workbook(filename)
    ws_source = wb.active
    current_date = today.strftime("%d.%m.%Y")
    ws_source.cell(row=3, column=1).value = current_date
    
    # Сохраняем с новой датой в имени
    base, ext = os.path.splitext(filename)
    new_filename = f"../{base} {current_date}{ext}"

    wb.save(new_filename)

    print(f"Файл {filename} обработан и сохранён как {new_filename}")
else:
    print(f"На {day_of_week} нет доступного для обработки файла.")
