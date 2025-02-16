# -*- coding: utf-8 -*-
import json
from datetime import datetime

from openpyxl import Workbook

current_date = datetime.now().date()  # Получение текущей даты
excel_filename = f"Список участков на {current_date}.xlsx"  # Создаем временный файл Excel

wb = Workbook()  # Создаем рабочую книгу Excel
ws = wb.active  # Создаем активную таблицу Excel

with open('rap_2024.json', 'r', encoding='utf-8') as json_file:  # Открываем JSON файл с данными
    list_plots = json.load(json_file)  # Используем функцию load для загрузки данных из файла

print(list_plots)  # Выводим список участков

ws.append(["Код участка", "Название участка"])  # Заголовки столбцов таблицы Excel с данными

for code, name in list_plots.items():  # Добавляем данные в таблицу Excel по коду участка и названию участка
    ws.append([code, name])  # Добавляем данные в таблицу с кодом участка и названием участка в столбце

wb.save(excel_filename)  # Сохраняем файл Excel в папку с данными
