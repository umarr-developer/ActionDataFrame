import pandas
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
        self.data.to_excel(f'{path}/{self.name}.xlsx', index=False, header=False)
