from datetime import datetime

import openpyxl
import pandas
from openpyxl.utils.dataframe import dataframe_to_rows


class AloneAct:
    name: str
    branches: list[str]
    filter_column: str
    columns: list[str]
    data: pandas.DataFrame
    signature: bool
    count: int

    def __init__(self, name: str, branches: list[str], filter_column: str, columns: list[str],
                 all_data: pandas.DataFrame, signature: bool = True):
        self.name = name
        self.branches = branches
        self.filter_column = filter_column
        self.columns = columns
        self.data = self.filter_act(all_data)
        self.signature = signature
        self.count = len(self.data)

    def filter_act(self, all_data: pandas.DataFrame):
        """
        Фильтрует из общих данных нужные столбцы и поля
        """
        branches = self.branches
        data_ = all_data.filter(items=self.columns)
        return data_.query(f'{self.filter_column} in @branches')

    def get_count(self) -> int:
        """
        Возвращает кол-во строк
        """
        return len(self.data.index)

    def write_act(self, path: str):
        """
        Записывает файлы с даннымии в указанныую папку
        """
        format_act = AloneActFormat(self)
        format_act.write_to_excel(path)


class AloneActFormat:
    template_folder = 'templates'
    template_name = 'AloneActTemplate.xlsm'
    sheet_name = 'Лист1'
    start_pos = 7

    def __init__(self, act: AloneAct):
        self.act = act
        self.workbook = openpyxl.load_workbook(f'{self.template_folder}/{self.template_name}')
        self.sheet = self.workbook[self.sheet_name]

    def set_print_area(self, rows_count: int, alt_index: int, signature: bool):
        index = self.start_pos + rows_count

        self.sheet.page_setup.paperSize = self.sheet.PAPERSIZE_A4
        self.sheet.print_area = f'A1:F{index + alt_index}'

    def set_values(self, rows_count: int, signature: bool):
        alt_index = 0

        date = datetime.now()
        self.sheet['F3'] = f'{date.day}.{date.month}.{date.year}'
        self.sheet['A6'] = f'{self.sheet["A6"].value} {self.act.name}'

        if signature:
            index = self.start_pos + rows_count + 3
            alt_index += 10
            if index % 79 > 79 - 8:
                index += 10
                alt_index += 10

            self.sheet[f'B{index}'] = 'Передал сотрудник ОЭиПК:'
            self.sheet[f'B{index + 3}'] = 'Передал: '
            self.sheet[f'B{index + 4}'] = 'Выездной сотрудник ОЛ '
            self.sheet[f'B{index + 5}'] = '(ФИО и подпись)'

            self.sheet[f'E{index}'] = 'Принял Сотрудник ОЛ:'
            self.sheet[f'E{index + 1}'] = '(ФИО и подпись)'
            self.sheet[f'E{index + 3}'] = 'Принял: '
            self.sheet[f'E{index + 4}'] = 'Нач ОО/зав с/к'
            self.sheet[f'E{index + 5}'] = '(ФИО и подпись)'

        self.set_print_area(rows_count, alt_index, signature)

    def write_to_excel(self, path: str):
        self.set_values(self.act.count, self.act.signature)

        for index, row in enumerate(dataframe_to_rows(self.act.data, index=False, header=False), start=1):
            row.insert(0, index)
            for index_, value in enumerate(row):
                self.sheet.cell(self.start_pos + index, index_ + 1, value)

        self.workbook.save(f'{path}/{self.act.name}.xlsx')
