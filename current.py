import openpyxl
import tools


# 将worksheet中index列中的内容按照指定的分隔符分割, 并将结果分割成number列
def text_to_columns(index: int, delimiter: str, number: int, worksheet):
    # 首先确定填充范围, 也是从2-max_rows-1
    max_row = ws.max_row

    # 然后需要准备空行, 注意将结果分割成number列, 即在index列后插入number-1列, 所以先插入列
    worksheet.insert_cols(index + 1, number - 1)

    # 然后把列头填充一下, 从index 填充到index+number-1, 一共 number列, 这里因为是range所以正好替我们减了1
    for i in range(index, index + number):
        ws.cell(row=1, column=i, value=str(i - index + 1) + '级科目')

    # 之后每一行进行分割, 然后向同一行填充, 这里注意如果分割数量多, 需要控制不要填充出number的范围
    for i in range(2, max_row):
        # 分割index列的单元格的字符串
        split_cell = ws.cell(row=i, column=index).value.split(delimiter)

        # 如果分割的长度大于number, 只填充到number为止
        if len(split_cell) > number:
            # 从(i, index+j)的单元格开始横向填充number数量的三个单元格
            for j in range(0, number):
                ws.cell(row=i, column=index + j, value=split_cell[j].strip())
        # 如果分割的长度小于等于number, 就直接填充即可
        else:
            j = 0
            for each_content in split_cell:
                ws.cell(row=i, column=index + j, value=each_content.strip())
                j = j + 1

    return worksheet


# 有了分列函数, 再来编写最终的处理函数
def process_worksheet(worksheet):
    # 填充5-6列
    tools.fill_column(5, worksheet)
    tools.fill_column(6, worksheet)

    # 删除列, 注意先删除右边的, 这样要删除的列号不会变化
    worksheet.delete_cols(15, 18)
    worksheet.delete_cols(10, 3)
    worksheet.delete_cols(8)
    worksheet.delete_cols(1, 4)

    # 分列, 注意此时分的是第四列
    text_to_columns(4, '-', 3, worksheet)

    return worksheet


if __name__ == '__main__':
    wb = openpyxl.load_workbook('kisdocument.xlsx')

    ws = wb.active

    process_worksheet(ws)

    ws.parent.save("result.xlsx")
