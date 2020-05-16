import os

from openpyxl import load_workbook


def excel_file_list(path: str) -> list:
    result_list = []

    # path : 'd:\\4月报表'

    # 对根路径中的每个文件记录进行迭代
    for each_entry in os.scandir(path):
        # 如果文件记录是一个文件夹/目录,
        if each_entry.is_dir():
            result_list += excel_file_list(each_entry.path)

        # 如果文件记录是一个普通文件,判断一下是不是excel文件,然后将其加入result_list
        else:
            if each_entry.path.endswith('.xlsx'):
                result_list.append(each_entry.path)

    return result_list


def excel_file_list_iter(path: str) -> list:
    # 创建两个数据结构
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

    # stack 已经为空
    return result_list


# 要读取所有文件的A3单元格

if __name__ == '__main__':

    for each_path in excel_file_list_iter('d:\\4月报表'):
        workbook = load_workbook(each_path)
        print(workbook.active['A3'].value)
        workbook.close()

# IDE
# 操作系统与文件系统
# 库与文档
# 递归与迭代
# 数据结构: 数组, 背包, 栈, 队列
