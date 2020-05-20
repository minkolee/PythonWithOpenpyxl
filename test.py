import openpyxl
import decimal


wb = openpyxl.load_workbook('test.xlsx')

wb.active['F14'] = str( decimal.Decimal('4.2') - decimal.Decimal('4.17'))

wb.save('test.xlsx')


