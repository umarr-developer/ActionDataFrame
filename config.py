from dataclasses import dataclass

from yaml import load, SafeLoader


@dataclass
class Config:
    filename: str
    filter_column: str
    columns: list[str]
    output: str
    branches: dict


def load_config(path: str, config_file: str) -> Config:
    with open(f'{path}/{config_file}', 'r', encoding='utf-8') as file:
        data = load(file, SafeLoader)
    return Config(
        filename=data['filename'],
        filter_column=data['filter_column'],
        columns=data['columns'],
        output=data['output'],
        branches=data['branches'])
