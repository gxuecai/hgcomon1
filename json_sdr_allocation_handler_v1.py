import json
import re

# ----------------------------Define----------------------------
class LteNR_ca_combo_json:
    '''
    [0] ".tech = RFDEVICE_TRX_STD_NR5G_SUB6",
    [1] ".wtr_band = SDR865_SUB6_RX_BAND41",
    [2] ".tx_pll = TRUE",
    [3] ".num_dl_layers = 4",
    [4] ".rxpll = SDR865_RX_PLL_RXPLL0",
    [5] ".main_dl_pair = SDR865_DL_PAIR_RX0",
    [6] ".mimo_dl_pair = SDR865_DL_PAIR_RX1",
    [7] ".main_prx_port_bitmask = SDR865_RX_PORT_PRX7",
    [8] ".main_drx_port_bitmask = SDR865_RX_PORT_DRX7",
    [9] ".mimo_prx_port_bitmask = SDR865_RX_PORT_PRX0",
    [10] ".mimo_drx_port_bitmask = SDR865_RX_PORT_DRX0",
    [11] ".lo_div = 2",
    [12] ".ilna_split_rxpll_bitmask = 0"
    '''
    def __init__(self,combos):
        self.combos = combos
        self.dl_ca_list = []
        self.ul_ca_list = []
        self.dl_band_num = 0
        self.band_list=[]
        self.get_ca_list(combos)
        self.get_dl_ca_band_list()
        # self.print_ca_info()

    def get_ca_list(self, combo_list):
        self.dl_band_num = len(combo_list)
        for band_item_json in combo_list:
            self.parse_band_info(band_item_json)

    def parse_band_info(self,band_item_json):
        bandenum_re = re.search(r'BAND([0-9]+)', band_item_json[1])
        tech_re = re.search(r'NR5G', band_item_json[0])
        if tech_re: # NR
            bandenum = 'N'+bandenum_re.group(1)
        else: # LTE
            bandenum = bandenum_re.group(1)

        ant_num = band_item_json[3][-1]
        self.dl_ca_list.append((bandenum,ant_num))

        if re.search(r'TRUE', band_item_json[2]):
            self.ul_ca_list.append(bandenum)

    def print_ca_info(self):
        print('dl_ca_list: ',self.dl_ca_list,' ul_ca_list:',self.ul_ca_list,' ca band number: ', self.dl_band_num)

    def get_dl_ca_band_list(self):
        self.dl_ca_list.sort()
        self.band_list = [item[0] for item in self.dl_ca_list]

def lte_nr_combo_nrx_handle(rx_num, combos_nrx_dict, keys_str_n, output_nrx_combo_json):
    '''
    :param rx_num: 1~5
    :param combos_nrx_dict: the dict content of rev'0x03' : dict_keys(['sdr865_ca_band_info_1rx_tbl_list_rev_0x03', 'sdr865_ca_band_info_2rx_tbl_list_rev_0x03',...
    :param keys_str_n: the corresponding key of nrx combo, such as sdr865_ca_band_info_1rx_tbl_list_rev_0x03
    :param output_nrx_combo_json: output combo
    :return: NULL
    '''
    re_str = ['0rx','1rx','2rx','3rx','4rx','5rx']
    if re.search(re_str[rx_num], keys_str_n): # Check if the key string contains 'nrx'
        print("\n>>> Matched %s tbl list" % re_str[rx_num])
        tbl_card_list_nrx_dict=combos_nrx_dict[keys_str_n] # get the dict content of nrx 'sdr865_ca_band_info_1rx_tbl_list_rev_0x03'
        len_tbl_card_list_nrx = len(tbl_card_list_nrx_dict) # As there may be several cards, print the length of nrx_dict
        print('The card numbers in %s combo: ' %re_str[rx_num], len_tbl_card_list_nrx)
        keys_tbl_card_list_nrx=tbl_card_list_nrx_dict.keys() # get the card keys in the nrx dict, dict_keys(['sdr865_ca_band_info_1rx_tbl_card0_rev_0x03', 'sdr865_ca_band_info_1rx_tbl_card2_rev_0x03'...
        print(keys_tbl_card_list_nrx, '\nCombo number in each card: ')

        for key_tbl_card_nrx in keys_tbl_card_list_nrx: # for the card dict
            tbl_card_dict_i= tbl_card_list_nrx_dict[key_tbl_card_nrx] # get the dict content of nrx card, 'sdr865_ca_band_info_1rx_tbl_card0_rev_0x03'
            combo_list = tbl_card_dict_i["list"] # get the list combos in card dict
            print(len(combo_list),end=' ') # This print end of ' ' instead of \n

            for combo_i in combo_list:
                output_nrx_combo_json.append(LteNR_ca_combo_json(combo_i))
    else:
        print("NO Matched %s tbl list" % re_str[rx_num])

    print('\nTotal combos of %s: ' % re_str[rx_num], len(output_nrx_combo_json))

# ---------------------------Test------------------------------


# ---------------------------Main--------------------------------
file_read = open(r"C:\CODE\MPSS.HI.1.0.c8-00198\modem_proc\rf\rfdevice_sdr865\common\etc\Storage\RF_SW\sdr865_default_ca_combo_allocations.json","rb")
sdr865_ca_combo_allocation = json.load(file_read) # json load file to dict variable

assert isinstance(sdr865_ca_combo_allocation, dict)

rev_list = sdr865_ca_combo_allocation.keys()
print('Revision List: ', rev_list)
for rev_list_i in rev_list:
    combos_rx_num_dict = sdr865_ca_combo_allocation[rev_list_i]
    print("Max combos RX number: ", len(combos_rx_num_dict))
    nrx_tbl_list = combos_rx_num_dict.keys();
    print('Tbl list:' , list(nrx_tbl_list))

    # 1RX case
    lte_nr_combo_json_1rx = []
    lte_nr_combo_nrx_handle(1,combos_rx_num_dict,list(nrx_tbl_list)[0], lte_nr_combo_json_1rx)
    '''
    if re.search(r'1rx', list(nrx_tbl_list)[0]):
        print("Matched 1RX tbl list")
        tbl_card_list_1rx_dict=combos_rx_num_dict[list(nrx_tbl_list)[0]]
        len_tbl_card_list_1rx = len(tbl_card_list_1rx_dict)
        print(len_tbl_card_list_1rx)
        keys_tbl_card_list_1rx=tbl_card_list_1rx_dict.keys()
        print(keys_tbl_card_list_1rx)

        for key_tbl_card_1rx in keys_tbl_card_list_1rx:
            tbl_card_dict_i= tbl_card_list_1rx_dict[key_tbl_card_1rx]
            combo_list = tbl_card_dict_i["list"]
            print(len(combo_list))

            for combo_i in combo_list:
                lte_nr_combo_json_1rx.append(LteNR_ca_combo_json(combo_i))
    else:
        print("Not matched 1RX tbl list")
    '''
    # 2RX case
    lte_nr_combo_json_2rx = []
    lte_nr_combo_nrx_handle(2, combos_rx_num_dict, list(nrx_tbl_list)[1], lte_nr_combo_json_2rx)

    '''
        if re.search(r'2rx', list(nrx_tbl_list)[1]):
        print("Matched 2RX tbl list")
        tbl_card_list_2rx_dict=combos_rx_num_dict[list(nrx_tbl_list)[1]]
        len_tbl_card_list_2rx = len(tbl_card_list_2rx_dict)
        print(len_tbl_card_list_2rx)
        keys_tbl_card_list_2rx=tbl_card_list_2rx_dict.keys()
        print(keys_tbl_card_list_2rx)

        for key_tbl_card_2rx in keys_tbl_card_list_2rx:
            tbl_card_dict_i= tbl_card_list_2rx_dict[key_tbl_card_2rx]
            combo_list = tbl_card_dict_i["list"]
            print(len(combo_list))

            for combo_i in combo_list:
                lte_nr_combo_json_2rx.append(LteNR_ca_combo_json(combo_i))
        else:
            print("Not matched 2RX tbl list")

    '''
    # 3RX case
    lte_nr_combo_json_3rx = []
    lte_nr_combo_nrx_handle(3, combos_rx_num_dict, list(nrx_tbl_list)[2], lte_nr_combo_json_3rx)

    '''
        if re.search(r'3rx', list(nrx_tbl_list)[2]):
        print("Matched 3RX tbl list")
        tbl_card_list_3rx_dict=combos_rx_num_dict[list(nrx_tbl_list)[2]]
        len_tbl_card_list_3rx = len(tbl_card_list_3rx_dict)
        print(len_tbl_card_list_3rx)
        keys_tbl_card_list_3rx=tbl_card_list_3rx_dict.keys()
        print(keys_tbl_card_list_3rx)

        for key_tbl_card_3rx in keys_tbl_card_list_3rx:
            tbl_card_dict_i= tbl_card_list_3rx_dict[key_tbl_card_3rx]
            combo_list = tbl_card_dict_i["list"]
            print(len(combo_list))

            for combo_i in combo_list:
                lte_nr_combo_json_3rx.append(LteNR_ca_combo_json(combo_i))
        else:
            print("Not matched 3RX tbl list")
    '''
    # 4RX case
    lte_nr_combo_json_4rx = []
    lte_nr_combo_nrx_handle(4, combos_rx_num_dict, list(nrx_tbl_list)[3], lte_nr_combo_json_4rx)
    # 5RX case
    lte_nr_combo_json_5rx = []
    lte_nr_combo_nrx_handle(5, combos_rx_num_dict, list(nrx_tbl_list)[4], lte_nr_combo_json_5rx)

file_read.close()


