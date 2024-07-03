from datetime import datetime

import openpyxl
import pandas
from openpyxl.utils.dataframe import dataframe_to_rows

from alone_act import AloneAct


class GeneralAct:
    name: str
    acts: list[AloneAct]
    data: pandas.DataFrame

    def __init__(self, name: str):
        self.name = name
        self.acts = []

    def add_act(self, act: AloneAct):
        """
        Добавляет запись об актах
        """
        self.acts.append(act)

    def act_to_data(self):
        data_ = {'column0': [], 'column1': []}
        for act in self.acts:
            data_['column0'].append(act.name)
            data_['column1'].append(act.get_count())
        self.data = pandas.DataFrame(data_)

    def write_act(self, path: str):
        """
        Записывает файлы с даннымии в указанныую папку
        """
        self.act_to_data()
        format_act = GeneralActFormat(self)
        format_act.write_to_excel(path)
        # self.data.to_excel(f'{path}/{self.name}.xlsx', index=False, header=False)


class GeneralActFormat:
    template_folder = 'templates'
    template_name = 'GeneralActTemplate.xlsm'
    sheet_name = 'Лист1'
    start_pos = 10

    def __init__(self, act: GeneralAct):
        self.act = act
        self.workbook = openpyxl.load_workbook(f'{self.template_folder}/{self.template_name}')
        self.sheet = self.workbook[self.sheet_name]

    def set_print_area(self, rows_count: int, alt_index: int):
        index = self.start_pos + rows_count

        self.sheet.page_setup.paperSize = self.sheet.PAPERSIZE_A4
        self.sheet.print_area = f'A1:C{index + alt_index}'

    def set_values(self):
        rows_count = len(self.act.acts)
        alt_index = 10

        self.set_print_area(rows_count, alt_index)

    def write_to_excel(self, path: str):
        self.set_values()

        for index, row in enumerate(dataframe_to_rows(self.act.data, index=False, header=False), start=1):
            row.insert(0, index)
            for index_, value in enumerate(row):
                self.sheet.cell(self.start_pos + index, index_ + 1, value)

        self.workbook.save(f'{path}/{self.act.name}.xlsx')
