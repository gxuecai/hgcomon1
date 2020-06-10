import re
import os
import xlsxwriter as XW
from tqdm import tqdm
from subset import get_subset_combo_string
import ca_combo_class

def compare_combo(rfc_combo, sdr_internal_combo):

    if rfc_combo.band_list != sdr_internal_combo.band_list:
        return False

    if rfc_combo.ul_band_list != sdr_internal_combo.ul_ca_list:
        return False

    rfc_ant_list = [item[1] for item in rfc_combo.dl_ca_list]
    sdr_ant_list = [item1[1] for item1 in sdr_internal_combo.dl_ca_list]

    ant_cmp = [(rfc_ant_list[item_i] <= sdr_ant_list[item_i]) for item_i in range(0,rfc_combo.dl_band_num)]

    if False in ant_cmp:
        return False
    else:
        return True

# load RFC combo
import xml_parser_rfc

# load sdr combo
import json_sdr_allocation_handler_v1

lte_nr_combo_nrx_list = [0,
                             json_sdr_allocation_handler_v1.lte_nr_combo_json_1rx,
                             json_sdr_allocation_handler_v1.lte_nr_combo_json_2rx,
                             json_sdr_allocation_handler_v1.lte_nr_combo_json_3rx,
                             json_sdr_allocation_handler_v1.lte_nr_combo_json_4rx,
                             json_sdr_allocation_handler_v1.lte_nr_combo_json_5rx]

# initial excel table
wb = XW.Workbook('CA_port_mapping.xlsx')
ws = wb.add_worksheet()

format1 = wb.add_format({'font_size': 10, 'font_name': 'Arial', 'bold': 1, 'font_color': 'white',
                         'fg_color': 'green',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})
format2 = wb.add_format({'font_size': 10, 'font_name': 'Arial', 'bold': 0, 'font_color': 'black',
                         'fg_color': '#F2F2F2',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})
format3 = wb.add_format({'font_size': 10, 'font_name': 'Arial', 'bold': 0, 'font_color': 'black',
                         'fg_color': 'white',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})

ws.set_column(0, 1, 25, format3)
ws.set_column(2, 2, 40, format2)
ws.set_column(3, 3, 40, format3)
ws.set_column(4, 4, 40, format2)
ws.set_column(5, 5, 40, format3)
ws.set_column(6, 6, 40, format2)
# ws.set_column(7,7,40,format3)
ws.freeze_panes(1, 0)

Heading_list = ['DL CA combos', 'CA combos', 'Band Idx 0', 'Band Idx 1', 'Band Idx 2', 'Band Idx 3',
                    'Band Idx 4']
ws.set_row(0, 40)
ws.write_row(0, 0, Heading_list, format1)
style_list = [format2, format3]
row_n = 1

#-----------------Subset------------------
all_combo_string = [combo_object.ca_string for combo_object in (xml_parser_rfc.lte_combo_list+xml_parser_rfc.nr_combo_list+xml_parser_rfc.endc_combo_list)]
all_subsets = []
for combo_string in all_combo_string:
    all_subsets += get_subset_combo_string(combo_string)

all_subsets_v = list(set(all_subsets))
all_subsets_v.sort()
all_subsets_object = []
fo = open("log.txt", "w")
for subset_string in all_subsets_v:
    all_subsets_object.append(ca_combo_class.LteNR_ca_combo(subset_string))
    fo.write(subset_string + '\n')
fo.close()
#-----------------------------------------
#combos_i = 0
matched_flag = 0
col0_style_index = 0

total_combo_number = len(all_subsets_object)
pbar = tqdm(total=total_combo_number) # 进度条

for rfc_combo in all_subsets_object:
    matched_flag = 0
    for sdr_internal_combo in lte_nr_combo_nrx_list[rfc_combo.dl_band_num]:

        if compare_combo(rfc_combo, sdr_internal_combo):
            matched_flag = 1
            dl_ca_string=''
            for band_i in range(0, rfc_combo.dl_band_num):
                band_s = rfc_combo.band_list[band_i]
                if band_s[0] == 'N':
                    dl_ca_string += band_s
                else:
                    dl_ca_string += ('B' + band_s)
                if band_i != (rfc_combo.dl_band_num - 1):
                    dl_ca_string += '+'
            ws.write(row_n, 0, dl_ca_string, style_list[col0_style_index])
            ws.write(row_n, 1, rfc_combo.ca_string, format3)

            col_j = 1
            for combo_info_i in sdr_internal_combo.combos:
                col_j += 1
                print_band_port_info = json_sdr_allocation_handler_v1.get_band_port_string(
                    combo_info_i)  # combo_info_i[1]+'\n'+combo_info_i[7]+'\n'+ combo_info_i[8]+'\n'+combo_info_i[9]+'\n'+ combo_info_i[10]
                ws.write(row_n, col_j, print_band_port_info, style_list[col_j % 2])
            ws.set_row(row_n, 70)
            row_n+=1
            # print(rfc_combo.ca_string)

    pbar.update(1)
    if matched_flag:
        col0_style_index += 1
        col0_style_index = (col0_style_index % 2)

wb.close() # save to xlsx file
