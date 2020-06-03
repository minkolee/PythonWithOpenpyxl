import tools
import openpyxl

# 导出的格式好在统一, 不用进行什么额外处理
# 这里要先打印一行出来展示一下是什么.
print(tools.open_xlsx_file('kisdocument.xlsx').max_row)
print(tools.open_xlsx_file('kisdocument.xlsx').max_column)


def fill_column(column: int, worksheet: str):
    ws = tools.open_xlsx_file(worksheet)

    # 由于文件格式统一, 可以直接设置起点行号是2, 结束的行号是 ws.max-row-1
    start = 2
    end = ws.max_row - 1
    print('start is {}, end is {}'.format(start, end))

    # 然后的思路:
    # 从起点到结尾 2 to end
    # 第一行必定是有值的, 从2开始, 向下寻找是None的单元格, 如果是None就填充. 如果越界或者不是None就停止
    # 停止之后, 要么已经越界, 要么停在下一张凭证的第一行, 更新当前的行数, 然后继续执行同样的循环

    current_row = start

    # 外层的循环, 保证不越界
    while current_row <= end:
        # 说明是凭证的第一行, 用一个变量保存一下这个格子的数据
        current_value = ws.cell(row=current_row, column=column).value
        print('当前凭证处理行号: {}, 凭证号是: {}'.format(current_row, current_value))

        # 从第一行之后的下一行进行填充, 一直填充到下一个不为None或者越界结束.
        # 下一行的位置
        next_row = current_row + 1

        # 从下一行开始, 满足是None而且没有越界, 复制current_value的值到这些格子里
        while (not ws.cell(row=next_row, column=column).value) and next_row <= end:
            print("正在填充的行号是: {}, 值是 {}".format(next_row,current_value))
            ws.cell(row=next_row, column=column, value=current_value)
            next_row = next_row + 1

        # 这个循环结束以后, next_row有两种情况, 第一个停在下一个不是None的位置, 第二, 超过end.
        print('一张凭证处理结束, 新行号: {}'.format(next_row))
        # 所以要进行判断, 大于end了, 直接break,说明处理完毕
        if next_row > end:
            print('行号: {} 超过 {}, 填充结束'.format(next_row,end))
            break
        # 如果没大于end, 此时的next_row指向下一条凭证的第一行
        else:
            # 将当前循环变量设置为next_row, 即从这一行开始继续执行相同的逻辑.
            current_row = next_row
            print('将从 {} 行开始处理下一张凭证'.format(current_row))
            print('-----------------------------------------------------------------------')


    # 如果执行到这里, 说明处理完毕
    # ws.parent就是ws的爸爸, 也就是工作簿对象
    ws.parent.save('result.xlsx')


if __name__ == '__main__':
    fill_column(6, 'kisdocument.xlsx')
