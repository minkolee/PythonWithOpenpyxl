# 用于生成给绿城制作模拟报表, 从凭证中分解出一级和二级科目的程序
import tools
import current
import openpyxl


# 填充凭证号填充, 然后检查*号, 只保留指定的*号对应的凭证
# file_name是导出的5月凭证
# 返回固定的中间文件名, 表示提取后的5月凭证数据
def clean_data(file_name: str):
    # 填充凭证号, 获取一个中间表
    ws = current.fill_column(6, file_name)

    print(ws)

    # 遍历, 将*号对应的凭证号放到一个字典中
    # greentown_entries字典用于记录所有的归属于绿城的凭证号.
    greentown_entries = {}

    # 避免ws.max_row变动, 一开始就取好
    row_number = ws.max_row

    for number in range(1, row_number):
        if ws.cell(row=number, column=7).value and ws.cell(row=number, column=7).value.startswith('*'):
            greentown_entries[ws.cell(row=number, column=6).value] = 1

    print(greentown_entries)

    new_wb = openpyxl.Workbook()

    new_ws = new_wb.active

    current_row = 1

    # 按凭证号复制内容
    for each_number in range(1, row_number):
        # 如果是对应的凭证, 就复制到新表中
        if ws.cell(row=each_number, column=6).value in greentown_entries.keys():
            # 复制凭证号,摘要,科目代码科目名称 借方,贷方到新的工作表中
            new_ws.cell(row=current_row, column=1, value=ws.cell(row=each_number, column=6).value)
            new_ws.cell(row=current_row, column=2, value=ws.cell(row=each_number, column=7).value)
            new_ws.cell(row=current_row, column=3, value=ws.cell(row=each_number, column=8).value)
            new_ws.cell(row=current_row, column=4, value=ws.cell(row=each_number, column=9).value)
            new_ws.cell(row=current_row, column=5, value=ws.cell(row=each_number, column=13).value)
            new_ws.cell(row=current_row, column=6, value=ws.cell(row=each_number, column=14).value)
            current_row += 1
    new_wb.save("提取后的凭证.xlsx")
    return '提取后的凭证.xlsx'


# 获取科目与科目代码字典的函数
def get_subjects(file_name: str):
    result_dict = {}

    ws = tools.open_xlsx_file('科目余额表.xlsx')

    row_max = ws.max_row

    for each in range(1, row_max + 1):
        result_dict[ws.cell(row=each, column=1).value.strip()] = ws.cell(row=each, column=2).value.strip()

    print("组装的字典是: {}".format(result_dict))

    return result_dict


# 向提取后的凭证写入一级科目和二级科目的功能.
def format_entry(file_name: str):
    # 组装字典
    subject_dict = get_subjects('科目余额表.xlsx')

    ws = tools.open_xlsx_file(file_name)
    row_max = ws.max_row

    for each in range(1, row_max + 1):
        ws.cell(row=each, column=7, value=ws.cell(row=each, column=4).value.split('-')[0].strip())
        ws.cell(row=each, column=8, value=subject_dict.get(ws.cell(row=each, column=3).value))

    ws.parent.save(file_name)

    return file_name


# 统计一级科目合计数
def count_result(file_name: str):
    result_dict = {}

    wb = openpyxl.load_workbook(file_name)

    ws = wb.active

    row_max = ws.max_row

    for i in range(1, row_max + 1):
        # 如果已经有这个键, 加上借方数字, 减去贷方数字
        if result_dict.get(ws.cell(row=i, column=7).value):
            result_dict[ws.cell(row=i, column=7).value] = result_dict.get(ws.cell(row=i, column=7).value) + ws.cell(
                row=i, column=5).value - ws.cell(row=i, column=6).value
        else:
            result_dict[ws.cell(row=i, column=7).value] = ws.cell(row=i, column=5).value - ws.cell(row=i,
                                                                                                   column=6).value
    print(result_dict)

    ws1 = wb.create_sheet('一级科目汇总')

    start = 1

    for key, value in result_dict.items():
        ws1.cell(row=start, column=1).value = key
        ws1.cell(row=start, column=2).value = value
        start = start+1

    wb.save(file_name)

if __name__ == '__main__':
    # 返回提取后的文件名
    file = clean_data('5月凭证.xlsx')

    format_entry(file)

    count_result(file)
