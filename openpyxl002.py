import openpyxl
import openpyxl001
import random


def combine_excels(path: str, filename: str):
    wb = openpyxl.Workbook()

    for i in range(50):
        ws = wb.create_sheet("第{}张表".format(i + 1))

        color_str = '{0:06X}'.format(random.randint(0, 0xFFFFFF))

        ws.sheet_properties.tabColor = color_str

        print("第{0:d}张表的颜色是 {1}".format(i + 1, color_str))

    # 处理一下字符串
    if len(filename == 0):
        filename = 'default.xlsx'

    if not filename.endswith('.xlsx'):
        filename = filename.split('.')[0] + '.xlsx'

    wb.remove(wb.active)

    wb.save(filename)


combine_excels('d:\\4月报表', 'combine')
