import pandas


class GeneralAct:
    name: str
    data_: dict

    def __init__(self, name: str):
        self.name = name
        self.data_ = {'column0': [], 'column1': []}

    def add_act(self, name: str, count: int):
        """
        Добавляет запись об актах
        """
        self.data_['column0'].append(name)
        self.data_['column1'].append(count)

    def write_act(self, path: str):
        """
        Записывает файлы с даннымии в указанныую папку
        """
        data = pandas.DataFrame(self.data_)
        data.to_excel(f'{path}/{self.name}.xlsx', index=False, header=False)
