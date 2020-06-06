import xlrd
import xlwt
import xlsxwriter as xw
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

wb = xw.Workbook('CA_combos.xlsx')
ws = wb.add_worksheet('CA_combos')

format1 = wb.add_format({'font_size': 10, 'font_color': 'black','fg_color':'#E6FFE6', 'bottom': 1, 'top':1, 'right':1, 'left':1, 'bold':1,
                         'align':'center', 'valign':'vcenter', 'font_name':'Arial','text_wrap':1})
Heading_list = ['DL CA combos', 'UL CA combos','Band Idx 0', 'Band Idx 1','Band Idx 2','Band Idx 3','Band Idx 4\nTest\nTEST']

ws.write_row('A1',Heading_list,format1)
#ws.set_row(0,20)
ws.set_column(0,1,20)
ws.set_column(2,7,40)
ws.freeze_panes(1,0)
#ws.set_row(0,cell_format = format1)

wb.close()


wb = xw.Workbook('CA_port.xlsx')
ws = wb.add_worksheet()

format1 = wb.add_format({'font_size': 10,'font_name' : 'Arial', 'bold': 1, 'font_color': 'white',
                         'fg_color': 'green',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})
format2 = wb.add_format({'font_size': 10,'font_name' : 'Arial', 'bold': 0, 'font_color': 'black',
                         'fg_color':'#F2F2F2',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})
format3 = wb.add_format({'font_size': 10,'font_name' : 'Arial', 'bold': 0, 'font_color': 'black',
                         'fg_color':'white',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})

format4 = wb.add_format({'hidden': 1})
ws.set_column(0,1,20,format3)
ws.set_column(2,2,40,format2)
ws.set_column(3,3,40,format3)
ws.set_column(4,4,40,format2)
ws.set_column(5,5,40,format3)
ws.set_column(6,6,40,format2)
ws.set_column(7,7,40,format3)
ws.freeze_panes(1,0)

Heading_list = ['DL CA combos', 'UL CA combos','Band Idx 0', 'Band Idx 1','Band Idx 2','Band Idx 3','Band Idx 4']
ws.set_row(0,40)
#ws.set_column('H:I',format4)
ws.write_row(0, 0, Heading_list, format1)


#ws.write_column(0,1,'',format2)

wb.close()



