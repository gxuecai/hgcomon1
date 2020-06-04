import xml_parser_rfc
import re
import os
import json_sdr_allocation_handler_v1
import xlwt

# parse the lte band number as B1+B3+B4 format
def parse_ca_bands_from_input(input_string):
    band_list_return=[]
    split_input= re.split(r'\+', input_string)
    for band_string in split_input:
        band = re.search(r'[BbNn]([0-9]+)', band_string)
        if band.group()[0] in ['N','n']: # NR band
            band_list_return.append('N'+band.group(1))
        else:
            band_list_return.append(band.group(1))
        band_list_return.sort()
    return band_list_return

# -----------------------------Test-------------------------------
#print(parse_ca_bands_from_input("B1+B3+B5"))

# -----------------------------Main-------------------------------
os.system("cls") # cmd window clear screen

isRFC = input("Searching combos in RFC or internal sdr allocation (1: RFC, 0/others: internal sdr allocation)>>")

user_search_combo=""
lte_nr_combo_nrx_list = [0,
                         json_sdr_allocation_handler_v1.lte_nr_combo_json_1rx,
                         json_sdr_allocation_handler_v1.lte_nr_combo_json_2rx,
                         json_sdr_allocation_handler_v1.lte_nr_combo_json_3rx,
                         json_sdr_allocation_handler_v1.lte_nr_combo_json_4rx,
                         json_sdr_allocation_handler_v1.lte_nr_combo_json_5rx]
# ----------Excel handler-----------

wb = xlwt.Workbook()
print(type(wb)); assert isinstance(wb, xlwt.Workbook)
ws = wb.add_sheet('Search CA', cell_overwrite_ok = True)
print(type(ws)); assert isinstance(ws, xlwt.Worksheet)

style = xlwt.XFStyle()
font = xlwt.Font()
font.name = 'Arial'
font.bold = True
font.colour_index = xlwt.Style.colour_map['red']
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']
style.font = font
style.alignment.wrap = style.alignment.WRAP_AT_RIGHT # 自动换行
style.alignment.vert = style.alignment.VERT_CENTER
style.alignment.horz = style.alignment.HORZ_CENTER
#style.pattern = pattern

style2 = xlwt.XFStyle()
pattern2 = xlwt.Pattern()
pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
pattern2.pattern_fore_colour = xlwt.Style.colour_map['white']
style2.alignment.wrap = style2.alignment.WRAP_AT_RIGHT
style2.alignment.vert = style2.alignment.VERT_CENTER
#style2.pattern = pattern2

style_list = [style, style2]
#-----------------------------------
while 1:
    user_search_combo = input("Please input search combos as the format B1+B2+B4+N66 >>")
    if(user_search_combo == 'exit'):
        break
    input_band_combo = parse_ca_bands_from_input(user_search_combo)
    print('-------------------------------------------------- \nYour input band list: ', input_band_combo)

    if isRFC == '1':
        # find matched ca combos from RFC parser ca list
        for ca_combos_i in xml_parser_rfc.lte_combo_list+xml_parser_rfc.nr_combo_list+xml_parser_rfc.endc_combo_list:
            if input_band_combo == ca_combos_i.band_list:
                print(ca_combos_i.ca_string)
    else:
        combo_band_number = len(input_band_combo)
        row_n = 1
        ws.write(0, 0, 'DLCA combo', style)
        ws.write(0, 1, 'ULCA combo', style)
        ws.write(0, 2, 'Band Index 1', style)
        ws.write(0, 3, 'Band Index 2', style)
        ws.write(0, 4, 'Band Index 3', style)
        ws.write(0, 5, 'Band Index 4', style)
        ws.write(0, 6, 'Band Index 5', style)

        for ca_combos_i in lte_nr_combo_nrx_list[combo_band_number]:
            if input_band_combo == ca_combos_i.band_list:
                print(ca_combos_i.dl_ca_list)
                ws.write(row_n, 0, str(ca_combos_i.dl_ca_list), style2)
                ws.write(row_n, 1, str(ca_combos_i.ul_ca_list), style2)

                col_j = 1
                for combo_info_i in ca_combos_i.combos:
                    col_j += 1
                    print_band_port_info = combo_info_i[1]+'\n'+combo_info_i[7]+'\n'+ combo_info_i[8]+'\n'+combo_info_i[9]+'\n'+ combo_info_i[10]
                    ws.write(row_n, col_j, print_band_port_info, style2)
                row_n += 1
        for col_n in range(2,col_j+1):
            ws.col(col_n).width = 256*50
        wb.save('CA_port.xls')


    print('-------------------------------------------------- \n')

os.system("pause") # cmd window pause screen
