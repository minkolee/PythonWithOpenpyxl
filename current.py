import openpyxl
import tools
# 一开始不要引入颜色
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Color, NamedStyle, Side

# 要操作的表格
ws = tools.open_xlsx_file('sales.xlsx')
# 样式已经设置好的表格
target_ws = ws.parent['销售清单']

# 设置表头
ws.merge_cells('a1:f1')

# 字体与颜色

color1 = Color(theme=6, tint=-0.249977111117893)
font = Font(name="Microsoft YaHei UI", size=34, b=True, color=color1)
a1 = ws['a1']

alignment = Alignment(horizontal='center', vertical='center')

a1.font = font
a1.alignment = alignment

# -------------------行高与列宽-----------------------
# 这里先用两个循环把行高调成一致.

row_max = target_ws.max_row
col_max = target_ws.max_column

for i in range(1, row_max + 1):
    ws.row_dimensions[i].height = target_ws.row_dimensions[i].height
for i in range(1, col_max + 1):
    ws.column_dimensions[chr(64 + i)].width = target_ws.column_dimensions[chr(64 + i)].width

# 第二行, 由于6个格子一样, 所以可以考虑用模板样式

style_for_row2 = NamedStyle(name='header')
style_for_row2.font = Font(name='Calibri', size=16, color='FFFFFF')
style_for_row2.alignment = Alignment(horizontal='center', vertical='center')
style_for_row2.fill = PatternFill('solid', fgColor=color1)

for each_cell in ws[2]:
    each_cell.style = style_for_row2

# 开始表格部分, 这次也一样, 颜色把握不住就来打印一下看看

style_for_body = NamedStyle(name='body')
style_for_body.font = Font(name='Calibri')
style_for_body.alignment = Alignment(horizontal='center', vertical='center')

color2 = Color(theme=2, tint=-0.1499984740745262)

style_for_body.fill = PatternFill("solid", fgColor=color2)

print(target_ws['a3'].border.left)

style_for_body.border = Border(left=Side(border_style='thin', color='FF000000'),
                               right=Side(border_style='thin', color='FF000000'),
                               bottom=Side(border_style='thin', color='FF000000'),
                               top=Side(border_style='thin', color='FF000000')
                               )

for i in range(3, row_max + 1):
    for j in range(1, col_max + 1):
        ws.cell(row=i, column=j).style = style_for_body

ws.parent.save('new.xlsx')
