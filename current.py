import openpyxl
import tools


def subtotal_single(key_column:int, value_column:int, file_name:str)->dict:
    wb = tools.open_xlsx_file(file_name)

    result = {}

    # 2 -> wb.max_row -1 是要处理的列
    for i in range(2, wb.max_row):

        # 键是否存在于字典中
        key = wb.cell(row=i,column=key_column).value

        if key in result:

            # 存在的话, 需要更新
            result[key] = result[key] + tools.transfer_to_decimal(wb.cell(row=i,column=value_column).value)

        #不存在的话, 直接设置
        else:
            result[key] = tools.transfer_to_decimal(wb.cell(row=i,column=value_column).value)

    return result

def subtotal_composite(key_column:int, value_column1:int, value_column2:int, file_name:str) -> dict:

    wb = tools.open_xlsx_file(file_name)

    result = {}
    # 2 -> wb.max_row -1 是要处理的列
    for i in range(2, wb.max_row):

        # 键是否存在于字典中
        key = wb.cell(row=i, column=key_column).value

        if key in result:
            # 如果存在, 要更新两个值
            result[key][wb.cell(row=1,column=value_column1).value] = result[key][wb.cell(row=1,column=value_column1).value] + tools.transfer_to_decimal(wb.cell(row=i, column=value_column1).value)
            result[key][wb.cell(row=1,column=value_column2).value] = result[key][wb.cell(row=1,column=value_column2).value] + tools.transfer_to_decimal(wb.cell(row=i, column=value_column2).value)

        # 不存在的话, 创建键和对应的嵌套字典, 初始值是0
        else:
            result[key] = {}
            result[key][wb.cell(row=1,column=value_column1).value] = tools.transfer_to_decimal(wb.cell(row=i, column=value_column1).value)
            result[key][wb.cell(row=1,column=value_column2).value] = tools.transfer_to_decimal(wb.cell(row=i, column=value_column2).value)

    return result


if __name__ == '__main__':

    result = subtotal_single(4,7,'data.xlsx')

    # 为了方便, 可以写一个函数将内容输出

    wb = openpyxl.Workbook()
    ws = wb.active

    ws.cell(1,1,"科目")
    ws.cell(1,2,"借方")

    start = 2

    for(k,v) in result.items():
        ws.cell(row=start, column=1, value=k)
        ws.cell(row=start, column=2, value=v)
        start = start+1

    wb.save('result.xlsx')
