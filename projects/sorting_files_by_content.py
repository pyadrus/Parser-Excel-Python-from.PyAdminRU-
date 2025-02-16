import shutil
from pathlib import Path

import win32com.client as win32
from docx import Document


def search_in_docx(file_path, search_term):
    """
    Поиск в документе
    :param file_path: путь к файлу
    :param search_term: искомое слово
    :return: True, если слово найдено, False - если нет
    """
    doc = Document(file_path)
    for paragraph in doc.paragraphs:
        if search_term in paragraph.text:
            return True
    return False


def search_in_doc(file_path, search_term):
    """
    Поиск в документе
    :param file_path: путь к файлу
    :param search_term: искомое слово
    :return: True, если слово найдено, иначе False
    """
    # Используем COM для открытия .doc файлов
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(str(file_path))  # Приводим file_path к строке
    content = doc.Content.Text
    doc.Close()
    word.Quit()
    return search_term in content


def move_file(file_path, destination_folder):
    """
    Перемещение файла
    :param file_path: путь к файлу
    :param destination_folder: папка для перемещения файла
    :return: None
    """
    destination_folder.mkdir(parents=True, exist_ok=True)  # Создаем папки, если их нет
    shutil.move(str(file_path), str(destination_folder / file_path.name))


def sort_files_by_year(source_folder, destination_folder, year):
    """
    Сортировка файлов по содержимому
    :param source_folder: папка с файлами
    :param destination_folder: папка для перемещения файлов
    :param year: год, по которому будет производиться поиск
    :return: None"""
    search_term = str(year)
    for file_path in source_folder.iterdir():

        print(file_path)

        if file_path.suffix == '.docx' and search_in_docx(file_path, search_term):
            move_file(file_path, destination_folder)
        elif file_path.suffix == '.doc' and search_in_doc(file_path, search_term):
            move_file(file_path, destination_folder)


def sorting_files_by_content():
    """
    Сортировка файлов по содержимому
    :return: None
    """
    source_folder = Path(
        r'C:\Users\zhvit\YandexDisk\ГУП ДНР им. А.Ф. Засядько\Приказы_Распоряжения_Протоколы_Пояснение')
    destination_folder = source_folder / '2018'

    sort_files_by_year(source_folder, destination_folder, '2018')


if __name__ == '__main__':
    sorting_files_by_content()
