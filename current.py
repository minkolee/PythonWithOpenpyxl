import openpyxl
import tools

ws = tools.open_xlsx_file('sample.xlsx')

ws.iter_rows()