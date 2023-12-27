from dataclasses import dataclass

from yaml import load, SafeLoader


@dataclass
class Config:
    filename: str
    filter_column: str
    columns: list[str]
    output0: str
    branches0: dict
    output1: str
    branches1: dict


def load_config(path: str, config_file: str) -> Config:
    with open(f'{path}/{config_file}', 'r', encoding='utf-8') as file:
        data = load(file, SafeLoader)
    return Config(
        filename=data['filename'],
        filter_column=data['filter_column'],
        columns=data['columns'],
        output0=data['output0'],
        branches0=data['branches0'],
        output1=data['output1'],
        branches1=data['branches1'])
