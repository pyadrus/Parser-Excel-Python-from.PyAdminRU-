import openpyxl as op


def counting_duplicate_records():
    """Функция открытия списочного состава и подсчитывания дубликатов"""
    list_prof = []

    wb = op.load_workbook('Списочный состав 20240531.xlsx')
    ws = wb.active
    list_gup = []
    for row in ws.iter_rows(min_row=6, max_row=876, min_col=1, max_col=4):
        row_data = [cell.value for cell in row]
        list_gup.append(row_data)
        print(row_data)
        list_prof.append(row_data)
    print(list_prof)

    duplicates = {}
    for item in list_prof:
        key = tuple(item)  # создаем уникальный ключ из списка
        if key in duplicates:
            duplicates[key] += 1
        else:
            duplicates[key] = 1

    # подсчитываем количество дубликатов
    duplicate_count = 0
    for key, count in duplicates.items():
        if count > 1:
            duplicate_count += count - 1

    print(f"Количество дубликатов: {duplicate_count}")


if __name__ == "__main__":
    counting_duplicate_records()
