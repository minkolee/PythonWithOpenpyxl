import openpyxl

wb = openpyxl.load_workbook('test.xlsx')

ws = wb.active

ws1 = wb.copy_worksheet(wb.active)

ws1.sheet_properties.tabColor = '00FFFF'

ws1.sheet_properties.tagname = "新工作表"

print(wb.active.title)

wb.save('new.xlsx')
