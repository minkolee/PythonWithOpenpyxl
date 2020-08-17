import random
import sales
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Color, NamedStyle, Side

def setStyle(generation_func, title_color='44777b', header_color='44777b', background='d9d9d9'):
    #---------------------#
    ws = generation_func()
    #---------------------#
    ws.merge_cells('a1:f1')
    color1 = Color(rgb=title_color)
    font = Font(name="Microsoft YaHei UI", size=34, b=True, color=color1)
    a1 = ws['a1']
    alignment = Alignment(horizontal='center', vertical='center')
    a1.font = font
    a1.alignment = alignment

    row_max = ws.max_row
    col_max = ws.max_column

    for i in range(1, row_max + 1):
        ws.row_dimensions[i].height = 49.5
    for i in range(1, col_max + 1):
        ws.column_dimensions[chr(64 + i)].width = 15

    color2 = Color(rgb=header_color)

    style_for_row2 = NamedStyle(name='header')
    style_for_row2.font = Font(name='Calibri', size=16, color='FFFFFF')
    style_for_row2.alignment = Alignment(horizontal='center', vertical='center')
    style_for_row2.fill = PatternFill('solid', fgColor=color2)

    for each_cell in ws[2]:
        each_cell.style = style_for_row2

    color3 = Color(rgb=background)

    style_for_body = NamedStyle(name='body')
    style_for_body.font = Font(name='Calibri')
    style_for_body.alignment = Alignment(horizontal='center', vertical='center')
    style_for_body.fill = PatternFill("solid", fgColor=color3)
    style_for_body.border = Border(left=Side(border_style='thin', color='FF000000'),
                                   right=Side(border_style='thin', color='FF000000'),
                                   bottom=Side(border_style='thin', color='FF000000'),
                                   top=Side(border_style='thin', color='FF000000')
                                   )
    for i in range(3, row_max + 1):
        for j in range(1, col_max + 1):
            ws.cell(row=i, column=j).style = style_for_body

    return ws


setStyle(sales.sales_list).parent.save("new.xlsx")