import sqlite3
from tkinter import *
from tkinter.filedialog import askopenfilename
import os
from openpyxl import load_workbook

table_name = "parsing"  # Имя таблицы в базе данных
file_database = "data.db"  # Имя файла базы данных


def opening_a_file():
    """Открытие файла Excel"""
    root = Tk()
    root.withdraw()
    filename = askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    return filename


def parsing_document(min_row, max_row, column) -> None:
    """
    Осуществляет парсинг данных из файла Excel и вставляет их в базу данных SQLite.

    Аргументы:
    :param min_row: Строка, с которой начинается считывание данных.
    :param max_row: Строка, с которой заканчивается считывание данных.
    :param column: Столбец, с которого начинается считывание данных.
    """
    filename = opening_a_file()  # Открываем выбор файла Excel для чтения данных
    workbook = load_workbook(filename=filename)  # Загружаем выбранный файл Excel
    sheet = workbook.active

    os.remove(file_database)  # Удаляем файл базы данных

    conn = sqlite3.connect(file_database)  # Создаем соединение с базой данных
    cursor = conn.cursor()
    # Создаем таблицу в базе данных, если она еще не существует
    cursor.execute(f"CREATE TABLE IF NOT EXISTS {table_name} (service_number)")
    # Считываем данные из колонки A и вставляем их в базу данных
    for row in sheet.iter_rows(min_row=int(min_row), max_row=int(max_row), values_only=True):
        service_number = str(row[int(column)])  # Преобразуем значение в строку
        # Проверяем, существует ли запись с таким табельным номером в базе данных
        cursor.execute(f"SELECT * FROM {table_name} WHERE service_number = ?", (service_number,))
        existing_row = cursor.fetchone()
        # Если запись с таким табельным номером не существует, вставляем данные в базу данных
        if existing_row is None:
            cursor.execute(f"INSERT INTO {table_name} VALUES (?)", (service_number,))
    # Удаляем повторы по табельному номеру
    cursor.execute(
        f"DELETE FROM {table_name} WHERE rowid NOT IN (SELECT min(rowid) FROM {table_name} GROUP BY service_number)")
    # Сохраняем изменения в базе данных и закрываем соединение
    conn.commit()
    conn.close()
