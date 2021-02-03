from openpyxl import load_workbook

search_scheme_title_page = {
    'Номер объекта:': 1,
    'Наименование:': 1,
    'Адрес:': 1,
    'Дата проведения работ:': 1,
    'Исполнитель 1:': 2
    }

search_scheme_EPU = {
    'Тип ЭПУ:': [6],
    'Тип нагрузки': [6],
    'Выходной ток (общий), А': [6],
    'Состояние ЭПУ': [5],
    'Количество групп АКБ:': [5],
    'Количество АКБ в группе:': [5],
    'Тип АКБ:': [5, 6, 7, 8],
    'Всего АКБ в данном ЭПУ': [5],
    'Сумма номинальных емкостей АКБ, Ач:': [5],
    'Вывод': [1, 4],
    'Время автономной работы ориентировочно': [3],
    'Расчетное время на АКБ:': [3]
    }

TEST_CELL = 'G17'   # Если в ячейки есть данные то лист ексел нужно считать (только для ЭПУ)

wb = load_workbook('./25001066368-2020-COMPLEX-1.xlsm')
# wb = load_workbook('./test.xlsx')
sh_names = wb.sheetnames

# Читаю титульный лист
# ws = wb[sh_names[0]]
# for index, row in enumerate(ws.iter_rows()):
#     val = row[0].value
#     if val is None:
#         continue
#     for key, shift in search_scheme_title_page.items():
#         if val == key:
#             print(f'{row[0].value} {row[shift].value}')
#             break

# Читаю очет ЭПУ
for sh_name in sh_names:
    if sh_name[:3] == 'ЭПУ':
        ws = wb[sh_name]
        print(f'==========={sh_name}============')
        if ws[TEST_CELL].value is None:
            continue
        for index, row in enumerate(ws.iter_rows()):
            val = row[0].value
            if val is None:
                continue
            for key, shifts in search_scheme_EPU.items():
                if val == key:
                    for shift in shifts:
                        print(f'{row[0].value} {row[shift].value}')
                    break

# print(sh_names)