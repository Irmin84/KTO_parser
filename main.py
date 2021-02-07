# -*- coding: utf-8 -*-

import os

import openpyxl
from openpyxl import load_workbook

from utils import time_track


class ParserKTO:
    # В значениях словарей уазаны номера столбцов, для считывания значений (начиная с 0 для екселя)
    search_scheme_title = {
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

    def __init__(self, list_files):

        # self.end_row_number = 1
        self.my_wb = None
        self.my_sheet = None
        self.list_files = list_files
        self.file_to_create = {
            'Name': "./Report_KTO_test.xlsx",
            'sheet.title': "Данные КТО",
            'end_row_number': 1
        }
        self._create_file()
        self.count_process = 0
        self.total_files = len(self.list_files)

    def run(self):
        for file in self.list_files:
            try:
                wb = load_workbook(file, data_only=True)
            except Exception as exc:
                log = open('log.txt', 'a', encoding='UTF-8')
                error_str = f'{file} --> {type(exc)} {exc} \n'
                log.write(error_str)
                log.close()
                continue

            if file.endswith('.xlsm'):
                self._checking_new_report(wb)
            elif file.endswith('.xlsx'):
                self._checking_old_report(wb, file)
            # elif file.endswith('.pdf'):
            #     https: // dev - gang.ru / article / rabota - s - pdf - failami - v - python - cztenie - i - razbor - 06
            #     mta2spn0 /

            self.count_process += 1
            print(f'Считано {self.count_process}, осталось считать {self.total_files - self.count_process} файлов.')

    def _checking_new_report(self, wb):
        title_data_for_write = []
        epu_data_for_write = []
        sh_names = wb.sheetnames
        # Читаю титульный лист
        ws = wb[sh_names[0]]
        for index, row in enumerate(ws.iter_rows()):
            val = row[0].value
            if val is None:
                continue
            for key, shift in ParserKTO.search_scheme_title.items():
                if val == key:
                    # print(f'{row[0].value} {row[shift].value}')
                    title_data_for_write.append(row[shift].value)
                    break
        # Читаю отчет ЭПУ (все листы)
        for sh_name in sh_names:
            if sh_name[:3] == 'ЭПУ':
                ws = wb[sh_name]
                if ws.sheet_state == 'hidden':
                    continue
                # if ws[ParserKTO.TEST_CELL].value is None:
                #     continue
                # print(f'==========={sh_name}============')
                for index, row in enumerate(ws.iter_rows()):
                    val = row[0].value
                    if val is None:
                        continue
                    for key, shifts in ParserKTO.search_scheme_EPU.items():
                        if val == key:
                            for shift in shifts:
                                # print(f'{row[0].value} {row[shift].value}')
                                epu_data_for_write.append(row[shift].value)
                            break
                write_row = title_data_for_write + epu_data_for_write
                self._write_row_to_file(data_for_write=write_row)
                self.epu_data_for_write = []


    def _checking_old_report(self, wb, file):
        # TEST_CELL = 'A13'  # Старый формат отчета А18 содерижит "Владелец объекта:"
        title_data_for_write = []
        epu_data_for_write = []
        sh_names = wb.sheetnames
        ws = wb[sh_names[0]]
        if ws['A13'] == "Владелец объекта:":
            test_data = []
            test_data.append(file)
            self._write_row_to_file(test_data)

    def _create_file(self):
        self.my_wb = openpyxl.Workbook()
        self.my_sheet = self.my_wb.active
        self.my_sheet.title = self.file_to_create["sheet.title"]
        self._write_row_to_file(data_for_write=ParserKTO.table_header)

    def _write_row_to_file(self, data_for_write):
        for column_number in range(len(data_for_write)):
            write_cell = self.my_sheet.cell(row=self.file_to_create['end_row_number'], column=column_number + 1)
            write_cell.value = data_for_write[column_number]
        self.file_to_create['end_row_number'] += 1
        self.my_wb.save(self.file_to_create['Name'])



@time_track
def get_list_of_file(dirpath):
    list_of_files = []
    try:
        file = open('links.txt', 'w', encoding='UTF-8')
        for dirpath, _, filenames in os.walk(dirpath):
            for filename in filenames:
                if filename.endswith('.xlsm'):
                    link = os.path.join(dirpath, filename)
                    list_of_files.append(link)
                    file.write(link + '\n')
                elif filename.endswith('.xlsx'):
                    link = os.path.join(dirpath, filename)
                    list_of_files.append(link)
                    file.write(link + '\n')
        file.close()
    except Exception as exc:
        log = open('log.txt', 'a', encoding='UTF-8')
        error_str = f'Error get file {filename} --> {type(exc)} {exc} \n'
        log.write(error_str)
        log.close()
    return list_of_files


@time_track
def main():
    # dir = os.path.normpath(r"C:\Users\m.tkachev\Desktop\python\KTO_parser\test dir")
    dir = os.path.normpath(r"\\ceph-msk\Юг ТО СС\2020\Элиста\Комплекстное ТО")
    list_of_files = get_list_of_file(dirpath=dir)
    print(f'Найдено файлов: {len(list_of_files)}')
    # list_of_files = ['./25001066368-2020-COMPLEX-1.xlsm', './25001100029-2020-COMPLEX-1.xlsm']
    # list_of_files = ['./25001100029-2020-COMPLEX-1.xlsm']
    parser = ParserKTO(list_files=list_of_files)
    parser.run()


if __name__ == '__main__':
    main()
