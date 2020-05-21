import openpyxl
import decimal
import re


def transfer_to_decimal(num) -> decimal.Decimal:
    # 如果是一个整数, 就直接进行转换
    if type(num) == int:
        return decimal.Decimal(num)

    # 如果是一个浮点数, 将其转换成2位小数的字符串表示, 然后使用字符串来创建Decimal对象
    elif type(num) == float:
        return decimal.Decimal('{0:.2f}'.format(num))

    # 如果是一个字符串, 需要判断其格式
    elif type(num) == str:

        # 判断字符串是不是一个十进制的小数的表示
        if re.match('[+-]?\\d+(\\.\\d+)?$', num):
            # 字符串是一个十进制的小数表示
            # 判断是否是整数, 如果是整数, 直接通过整数创建Decimal对象
            if num.find('.') == -1:
                return decimal.Decimal(num)

            # 不是整数的情况下
            else:
                split_num = num.split('.')

                if len(split_num[1]) > 2:
                    num_string = split_num[0] + '.' + split_num[1][0:2]
                    if int(split_num[1][2]) >= 5:
                        if num[0] == '-':
                            return decimal.Decimal(num_string) - decimal.Decimal('0.01')
                        else:
                            return decimal.Decimal(num_string) + decimal.Decimal('0.01')
                    else:
                        return decimal.Decimal(num_string)

                elif len(split_num[1]) == 2:
                    num_string = split_num[0] + '.' + split_num[1][0:2]
                    return decimal.Decimal(num_string)

                elif len(split_num[1]) == 1:
                    num_string = split_num[0] + '.' + split_num[1] + '0'
                    return decimal.Decimal(num_string)

                else:
                    raise AttributeError

        else:
            raise AttributeError

    # 不是上述三种类型
    else:
        raise AttributeError


if __name__ == '__main__':
    wb = openpyxl.load_workbook('amount.xlsx')

    ws = wb.active

    # 先使用decimal来计算
    sum_decimal = decimal.Decimal('0')
    for i in range(3,61):
        sum_decimal = sum_decimal + transfer_to_decimal(ws['C'+str(i)].value)

    print(sum_decimal)

    # 使用float类型来计算
    sum_float = 0
    for i in range(3,61):
        sum_float = sum_float + float(ws['C'+str(i)].value)

    print(sum_float)