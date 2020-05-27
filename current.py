import openpyxl
import datetime

wb = openpyxl.load_workbook('datatypes.xlsx')

ws = wb.active

print("{} 的数据类型是 {}".format(ws['A1'].value, type(ws['A1'].value)))
print("{} 的数据类型是 {}".format(ws['A2'].value, type(ws['A2'].value)))
print("{} 的数据类型是 {}".format(ws['A3'].value, type(ws['A3'].value)))
print("{} 的数据类型是 {}".format(ws['A4'].value, type(ws['A4'].value)))
print("{} 的数据类型是 {}".format(ws['A5'].value, type(ws['A5'].value)))
print("{} 的数据类型是 {}".format(ws['A6'].value, type(ws['A6'].value)))
print("{} 的数据类型是 {}".format(ws['A7'].value, type(ws['A7'].value)))

ws['b1'] = 12345
ws['b2'] = 123.45
ws['b3'] = '123.45'
ws['b4'] = "字符串"
ws['b5'] = datetime.datetime(1899,5,27)
ws['b6'] = '=A1+A'

wb.save('new.xlsx')

wb = openpyxl.load_workbook('new.xlsx')
print()
print("{} 的数据类型是 {}".format(ws['B1'].value, type(ws['B1'].value)))
print("{} 的数据类型是 {}".format(ws['B2'].value, type(ws['B2'].value)))
print("{} 的数据类型是 {}".format(ws['B3'].value, type(ws['B3'].value)))
print("{} 的数据类型是 {}".format(ws['B4'].value, type(ws['B4'].value)))
print("{} 的数据类型是 {}".format(ws['B5'].value, type(ws['B5'].value)))
print("{} 的数据类型是 {}".format(ws['B6'].value, type(ws['B6'].value)))
