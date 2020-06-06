import xml_parser_rfc
import re
import os
import json_sdr_allocation_handler_v1
import xlsxwriter as XW

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

wb = XW.Workbook('CA_port.xlsx')
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

ws.set_column(0,1,20,format3)
ws.set_column(2,2,40,format2)
ws.set_column(3,3,40,format3)
ws.set_column(4,4,40,format2)
ws.set_column(5,5,40,format3)
ws.set_column(6,6,40,format2)
#ws.set_column(7,7,40,format3)
ws.freeze_panes(1,0)

Heading_list = ['DL CA combos', 'UL CA combos','Band Idx 0', 'Band Idx 1','Band Idx 2','Band Idx 3','Band Idx 4']
ws.set_row(0,40)
ws.write_row(0, 0, Heading_list, format1)

style_list = [format2, format3]
row_n = 1

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

        for ca_combos_i in lte_nr_combo_nrx_list[combo_band_number]:
            if input_band_combo == ca_combos_i.band_list:
                print(ca_combos_i.dl_ca_list)
                ws.write(row_n, 0, str(ca_combos_i.dl_ca_list), format3)
                ws.write(row_n, 1, str(ca_combos_i.ul_ca_list), format3)

                col_j = 1
                for combo_info_i in ca_combos_i.combos:
                    col_j += 1
                    print_band_port_info = json_sdr_allocation_handler_v1.get_band_port_string(combo_info_i) # combo_info_i[1]+'\n'+combo_info_i[7]+'\n'+ combo_info_i[8]+'\n'+combo_info_i[9]+'\n'+ combo_info_i[10]
                    ws.write(row_n, col_j, print_band_port_info, style_list[col_j % 2])
                ws.set_row(row_n, 70)
                row_n += 1


    print('-------------------------------------------------- \n')

wb.close()

os.system("pause") # cmd window pause screen
