import openpyxl
import tools
import configparser


def cal_current_ratio(worksheet, config) -> str:
    total_liquid_asset = worksheet[config['ratio']['liquid_asset']].value

    total_liquid_liability = worksheet[config['ratio']['liquid_liability']].value

    return "{:.2f}%".format(total_liquid_asset / total_liquid_liability * 100)


def load_config(file_name: str = None) -> configparser.ConfigParser:
    config = configparser.ConfigParser()
    if not file_name:
        config.read('config.ini')
    else:
        config.read(file_name)

    return config


if __name__ == '__main__':
    ws = tools.open_xlsx_file('600734.xlsx')

    conf = load_config()
    print('流动比率是: {}'.format(cal_current_ratio(ws, conf)))
