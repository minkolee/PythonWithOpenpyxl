import openpyxl
import tools

def fill_column(col_nubmer: int, worksheet):
    # 第一行固定指向2
    current = 2

    # 获取最大行
    max_index = worksheet.max_row - 1

    print("---------------填充从 {} 行开始, 到 {} 行结束---------------".format(current, max_index))
    # While current不越界:
    while current <= max_index :

        # 如果是None
        if not worksheet.cell(row=current, column=col_nubmer).value:
            # 填充上一格数据
            worksheet.cell(current, col_nubmer, worksheet.cell(current - 1, col_nubmer).value)
            print('填充第 {} 行, 值为 {}'.format(current, worksheet.cell(current - 1, col_nubmer).value))

        # current移动1格
        current += 1
        print("----------current 移动到 {} -----------".format(current))

    print("填充结束")

    return worksheet

if __name__ == '__main__':

    wb = openpyxl.load_workbook('kisdocument.xlsx')

    ws =wb.active

    fill_column(5,ws)
    fill_column(6,ws)

    wb.save('new.xlsx')
