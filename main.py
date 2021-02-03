from openpyxl import load_workbook

search_scheme_title_page = {
    'Номер объекта:': 1,
    'Наименование:': 1,
    'Адрес:': 1,
    'Дата проведения работ:': 1,
    'Исполнитель 1:': 2
    }

search_scheme_EPU = {
    'Тип ЭПУ:': 7,
    'Тип нагрузки': 8,
    'Выходной ток (общий), А': 7,
    'Состояние ЭПУ': 6,
    'Количество групп АКБ:': 6,
    'Количество АКБ в группе:': 6,
    'Тип АКБ:': [ 6 ,7, 8, 9],
    'Всего АКБ в данном ЭПУ': 6,
    'Сумма номинальных емкостей АКБ, Ач:': 7,
    'Вывод': [2, 5],
    'Время автономной работы ориентировочно': 5,
    'Расчетное время на АКБ:': 5
    }

wb = load_workbook('./25001066368-2020-COMPLEX-1.xlsm')
# wb = load_workbook('./test.xlsx')
sh_names = wb.sheetnames
ws = wb[sh_names[0]]


# Читаю титульный лист
for index, row in enumerate(ws.iter_rows()):
    # print(index, row
    val = row[0].value
    if val is None:
        continue
    for key, shift in search_scheme_title_page.items():
        if val == key:
            print(f'{row[0].value} {row[shift].value}')
            break

