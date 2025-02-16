# -*- coding: utf-8 -*-
import os

import xlrd


# Функция для сканирования папки и выполнения задачи для каждого файла
def process_xls_files():
    # Получить текущую рабочую папку
    current_directory = os.getcwd()

    for filename in os.listdir(current_directory):
        if filename.endswith(".xls"):
            file_path = os.path.join(current_directory, filename)

            # Открыть файл для чтения
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)  # Предполагаем, что нужный лист - первый

            # Прочитать текст из ячейки P7
            cell_value = sheet.cell_value(6, 15)  # 6 - номер строки (нумерация с 0), 15 - номер столбца (P)

            # Закрыть файл
            workbook.release_resources()
            del workbook

            # Генерировать новое имя файла с инкрементом, чтобы избежать конфликтов
            new_filename = os.path.join(current_directory, f"{cell_value}.xls")
            count = 1
            while os.path.exists(new_filename):
                new_filename = os.path.join(current_directory, f"{cell_value}_{count}.xls")
                count += 1

            # Переименовать файл
            os.rename(file_path, new_filename)


if __name__ == "__main__":
    process_xls_files()
