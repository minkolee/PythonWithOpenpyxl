import openpyxl
import tools

ws = tools.open_xlsx_file('origin.xlsx')

ws.merge_cells('D9:E10')

ws.parent.save('merged.xlsx')

ws = tools.open_xlsx_file('merged.xlsx')

print(ws.cell(row=9, column=4).value)
print(ws.cell(row=9, column=5).value)
print(ws.cell(row=10, column=4).value)
print(ws.cell(row=10, column=5).value)
print("---------------------------------")
print(type(ws.cell(row=9, column=4).value))
print(type(ws.cell(row=9, column=5).value))
print(type(ws.cell(row=10, column=4).value))
print(type(ws.cell(row=10, column=5).value))

ws.unmerge_cells('D9:E10')

ws.parent.save("unmerged.xlsx")
