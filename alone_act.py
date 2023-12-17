import pandas


class AloneAct:
    name: str
    branchs: list[str]
    filter_column: str
    columns: list[str]
    data: pandas.DataFrame

    def __init__(self, name: str, branchs: list[str], filter_column: str, columns: list[str],
                 all_data: pandas.DataFrame):
        self.name = name
        self.branchs = branchs
        self.filter_column = filter_column
        self.columns = columns
        self.data = self.filter_act(all_data)

    def filter_act(self, all_data: pandas.DataFrame):
        """
        Фильтрует из общих данных нужные столбцы и поля
        """
        branchs = self.branchs
        data_ = all_data.filter(items=self.columns)
        return data_.query(f'{self.filter_column} in @branchs')

    def write_act(self, path: str):
        """
        Записывает файлы с даннымии в указанныую папку
        """
        self.data.to_excel(f'{path}/{self.name}.xlsx', index=False, header=False)
