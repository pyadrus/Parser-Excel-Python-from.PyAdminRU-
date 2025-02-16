# -*- coding: utf-8 -*-
from openpyxl import load_workbook

# Имя файла
filename = "Копия T191651144.xlsx"

wb = load_workbook(filename, data_only=True)
sheet = wb.active
first_column_d = sheet['D']
first_column_e = sheet['E']
first_column_aw = sheet['AW']
first_column_at = sheet['AT']


def profession():
    """Профессия"""
    profession_dic = []  # Создаем словарь
    for x in range(len(first_column_d)):
        # if not first_column_d[x].value:
        #     continue
        profession_dic.append(first_column_d[x].value)
    return profession_dic


def discharge():
    """Разряд"""
    discharge_dic = []  # Создаем словарь
    for x in range(len(first_column_e)):
        # if not first_column_e[x].value:
        #     continue
        discharge_dic.append(first_column_e[x].value)
    return discharge_dic


def simple_enterprise():
    """Простой предприятия"""
    simple_enterprise_dic = []  # Создаем словарь
    for x in range(len(first_column_aw)):
        # if not first_column_e[x].value:
        #     continue
        simple_enterprise_dic.append(first_column_aw[x].value)
    return simple_enterprise_dic


def on_vacation():
    """В отпуске"""
    on_vacation_dic = []  # Создаем словарь
    for x in range(len(first_column_at)):
        # if not first_column_e[x].value:
        #     continue
        on_vacation_dic.append(first_column_at[x].value)
    return on_vacation_dic


def combining_dictionaries():
    profession_dic = profession()
    discharge_dic = discharge()
    simple_enterprise_dic = simple_enterprise()
    on_vacation_dic = on_vacation()
    for p in zip(profession_dic, discharge_dic, simple_enterprise_dic, on_vacation_dic):
        print(p)


combining_dictionaries()

if __name__ == "__main__":
    profession()
    discharge()
