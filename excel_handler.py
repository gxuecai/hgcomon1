import xlrd
import xlwt
#from pyExcelerator import *


wb = xlwt.Workbook()
print(type(wb)); assert isinstance(wb, xlwt.Workbook)
ws = wb.add_sheet('TEST', cell_overwrite_ok = True)
print(type(ws)); assert isinstance(ws, xlwt.Worksheet)

style = xlwt.XFStyle()
font = xlwt.Font()
font.name = 'Arial'
font.bold = 1
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']

style.font = font
style.alignment.wrap = style.alignment.WRAP_AT_RIGHT # 自动换行
style.alignment.shri = style.alignment.SHRINK_TO_FIT
style.pattern = pattern

ws.write(0, 0,'ahjfkahfjahsagsdjah\ntesat',style)

# set column width
first_col = ws.col(0)
print(type(first_col)); assert isinstance(first_col, xlwt.Column)
first_col.width = 256*50

wb.save("Test.xls")


