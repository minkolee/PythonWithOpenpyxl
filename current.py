import openpyxl
import decimal
import re
import datetime


ws = openpyxl.load_workbook('datatypes.xlsx').active

print("{} 的数据类型是 {}".format(ws['A1'].value, type(ws['A1'].value)))
print("{} 的数据类型是 {}".format(ws['A2'].value, type(ws['A2'].value)))
print("{} 的数据类型是 {}".format(ws['A3'].value, type(ws['A3'].value)))
print("{} 的数据类型是 {}".format(ws['A4'].value, type(ws['A4'].value)))
print("{} 的数据类型是 {}".format(ws['A5'].value, type(ws['A5'].value)))

print("{} 的数据类型是 {}".format(ws['B1'].value.today(), type(ws['B1'].value)))
print("{} 的数据类型是 {}".format(ws['B2'].value, type(ws['B2'].value)))
print("{} 的数据类型是 {}".format(ws['B3'].value, type(ws['B3'].value)))
print("{} 的数据类型是 {}".format(ws['B4'].value, type(ws['B4'].value)))
print("{} 的数据类型是 {}".format(ws['B5'].value, type(ws['B5'].value)))
