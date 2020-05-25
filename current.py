import openpyxl
import datetime

wb = openpyxl.load_workbook('datatypes.xlsx')

ws = wb.active

print("{} 的数据类型是 {}".format(ws['A1'].value, type(ws['A1'].value)))
print("{} 的数据类型是 {}".format(ws['A2'].value, type(ws['A2'].value)))
print()
print("{} 的数据类型是 {}".format(ws['B1'].value, type(ws['B1'].value)))
print("{} 的数据类型是 {}".format(ws['B2'].value, type(ws['B2'].value)))
print("{} 的数据类型是 {}".format(ws['B3'].value, type(ws['B3'].value)))
print("{} 的数据类型是 {}".format(ws['B4'].value, type(ws['B4'].value)))
print("{} 的数据类型是 {}".format(ws['B5'].value, type(ws['B5'].value)))
