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

    def __init__(self, name: str, branches: list[str], filter_column: str, columns: list[str],
                 all_data: pandas.DataFrame):
        self.name = name
        self.branches = branches
        self.filter_column = filter_column
        self.columns = columns
        self.data = self.filter_act(all_data)

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
        # self.data.to_excel(f'{path}/{self.name}.xlsx', index=False, header=False)


class AloneActFormat:
    template_folder = 'templates'
    template_name = 'AloneActTemplate.xlsx'
    sheet_name = 'Лист1'

    def __init__(self, act: AloneAct):
        self.act = act
        self.workbook = openpyxl.load_workbook(f'{self.template_folder}/{self.template_name}')
        self.sheet = self.workbook[self.sheet_name]

    def write_to_excel(self, path: str):
        date = datetime.now()

        self.sheet['F3'] = f'{date.day}.{date.month}.{date.year}'
        self.sheet['A6'] = f'{self.sheet["A6"].value} {self.act.name}'

        for index, row in enumerate(dataframe_to_rows(self.act.data, index=False, header=False), start=1):
            row.insert(0, index)
            self.sheet.append(row)

        self.workbook.save(f'{path}/{self.act.name}.xlsx')
