import pandas

from alone_act import AloneAct
from config import Config, load_config
from general_act import GeneralAct
from datetime import datetime

def action(all_data: pandas.DataFrame, config: Config, branches: dict, output: str, signature: bool):
    general_act = GeneralAct('_Общий акт')
    for branch in branches:
        # DEBUG
        print(f'>> {branch}')
        
        act = AloneAct(name=branch,
                       branches=branches[branch],
                       filter_column=config.filter_column,
                       columns=config.columns,
                       all_data=all_data,
                       signature=signature)
        act.write_act(output)
        general_act.add_act(act)
    general_act.write_act(output)


def main():
    # DEBUG
    start = datetime.now()
    
    config = load_config('config', 'config.yml')
    all_data = pandas.read_excel(config.filename)
    action(all_data, config, config.branches0, config.output0, signature=True)
    action(all_data, config, config.branches1, config.output1, signature=False)
    
    # DEBUG
    print(datetime.now() - start)


if __name__ == '__main__':
    main()
