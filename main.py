from openpyxl import load_workbook
import openpyxl

# В значениях словарей уазаны номера столбцов, для считывания значений (начиная с 0 для екселя)
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
    'Время автономной работы ориентировочно:': [3],
    'Расчетное время на АКБ:': [3]
    }

table_header = [
    'Номер объекта',
    'Наименование',
    'Адрес',
    'Дата проведения работ',
    'Исполнитель 1',
    'Тип ЭПУ',
    'Тип нагрузки',
    'Выходной ток (общий), А',
    'Состояние ЭПУ',
    'Количество групп АКБ',
    'Количество АКБ в группе',
    'Тип АКБ (группа 1)',
    'Тип АКБ (группа 2)',
    'Тип АКБ (группа 3)',
    'Тип АКБ (группа 4)',
    'Всего АКБ в данном ЭПУ',
    'Сумма номинальных емкостей АКБ, А/ч:',
    'Вывод',
    'Количество (Вывод)',
    'Время автономной работы (ориентировочно)',
    'Расчетное время на АКБ:'
    ]

TEST_CELL = 'G17'   # Если в ячейки есть данные то лист ексел нужно считать (только для ЭПУ)
titl_data_for_write = []
epu_data_for_write = []
end_row_number = 1
my_wb = None
my_sheet = None

wb = load_workbook('./25001066368-2020-COMPLEX-1.xlsm', data_only=True)
# wb = load_workbook('./test.xlsx')
sh_names = wb.sheetnames


# Запись в файл ttps://pythononline.ru/question/chtenie-i-zapis-v-fayl-excel-s-ispolzovaniem-modulya-python-openpyxl
def write_row_to_file(data_for_write):
    global end_row_number, my_wb, my_sheet
    if my_wb is None:
        my_wb = openpyxl.Workbook()
        my_sheet = my_wb.active
        my_sheet.title = "Данные КТО"
        for column_number in range(len(table_header)):
            write_cell = my_sheet.cell(row=end_row_number, column=column_number + 1)
            write_cell.value = table_header[column_number]
        end_row_number += 1

    for column_number in range(len(data_for_write)):
        write_cell = my_sheet.cell(row=end_row_number, column=column_number + 1)
        write_cell.value = data_for_write[column_number]
    end_row_number += 1
    my_wb.save("./Book1.xlsx")

# Читаю титульный лист
ws = wb[sh_names[0]]
for index, row in enumerate(ws.iter_rows()):
    val = row[0].value
    if val is None:
        continue
    for key, shift in search_scheme_title_page.items():
        if val == key:
            # print(f'{row[0].value} {row[shift].value}')
            titl_data_for_write.append(row[shift].value)
            break

# Читаю отчет ЭПУ
for sh_name in sh_names:
    if sh_name[:3] == 'ЭПУ':
        ws = wb[sh_name]
        if ws[TEST_CELL].value is None:
            continue
        print(f'==========={sh_name}============')
        for index, row in enumerate(ws.iter_rows()):
            val = row[0].value
            if val is None:
                continue
            for key, shifts in search_scheme_EPU.items():
                if val == key:
                    for shift in shifts:
                        print(f'{row[0].value} {row[shift].value}')
                        epu_data_for_write.append(row[shift].value)
                    break
        write_row = titl_data_for_write + epu_data_for_write
        write_row_to_file(data_for_write=write_row)
        epu_data_for_write = []
titl_data_for_write = []


