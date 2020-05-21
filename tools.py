import os
import openpyxl
import random


# 返回一个目录内所有的.xlsx文件的路径
def excel_file_list(path: str) -> list:
    result_list = []

    for each_entry in os.scandir(path):
        # 如果文件记录是一个文件夹/目录,
        if each_entry.is_dir():
            result_list += excel_file_list(each_entry.path)

        else:
            if each_entry.path.endswith('.xlsx'):
                result_list.append(each_entry.path)

    return result_list


# 返回一个目录内所有的.xlsx文件的路径
def excel_file_list_iter(path: str) -> list:
    result_list = []

    stack = [path]

    while len(stack) != 0:
        current_dir = stack.pop()
        if os.path.isdir(current_dir):
            for each_entry in os.scandir(current_dir):
                if each_entry.is_dir():
                    stack.append(each_entry.path)
                else:
                    if each_entry.path.endswith('.xlsx'):
                        result_list.append(each_entry.path)

        else:
            if current_dir.path.endswith('.xlsx'):
                result_list.append(current_dir.path)

    return result_list


# 月度化一个工作簿
def monthlize(path: str):
    wb = openpyxl.load_workbook(path)

    for i in range(12):
        ws = wb.copy_worksheet(wb.active)
        ws.title = "{}月{}".format(i + 1, wb.active.title)
        ws.sheet_properties.tabColor = '{0:06X}'.format(random.randint(0, 0xFFFFFF))

    wb.remove(wb.active)
    wb.save(path.split('.')[0] + 'monthly.xlsx')
