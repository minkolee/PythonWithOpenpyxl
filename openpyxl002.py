import openpyxl
import random


def monthlize(path: str):
    wb = openpyxl.load_workbook(path)

    for i in range(12):
        ws = wb.copy_worksheet(wb.active)

        ws.title = "{}月{}".format(i + 1, wb.active.title)

        ws.sheet_properties.tabColor = '{0:06X}'.format(random.randint(0, 0xFFFFFF))

    wb.remove(wb.active)

    wb.save(path.split('.')[0] + 'monthly.xlsx')


if __name__ == '__main__':
    monthlize('sales.xlsx')


