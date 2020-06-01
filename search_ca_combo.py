import xml_parser_rfc
import re
import os

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

user_search_combo=""
while 1:
    user_search_combo = input("Please input search combos as the format B1+B2+B4 >>")
    if(user_search_combo == 'exit'):
        break
    input_band_combo = parse_ca_bands_from_input(user_search_combo)
    print('-------------------------------------------------- \nYour input band list: ', input_band_combo)

    # find matched ca combos from RFC parser ca list
    for ca_combos_i in xml_parser_rfc.lte_combo_list+xml_parser_rfc.nr_combo_list+xml_parser_rfc.endc_combo_list:
        if input_band_combo == ca_combos_i.band_list:
            print(ca_combos_i.ca_string)
    print('-------------------------------------------------- \n')

os.system("pause") # cmd window pause screen
