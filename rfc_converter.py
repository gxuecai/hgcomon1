import xml.etree.ElementTree as ET
import xlsxwriter as XW
import types
import re


def calc_col_width_by_str(write_str):
    write_str = str(write_str)
    str_list = re.split(r'\n', write_str)
    str_len_list = [len(item) for item in str_list]
    return max(str_len_list)

def adjust_concurrency_number_str_col_width(num_col_width):

    if num_col_width > 180:
        num_col_width -= 40
    elif num_col_width > 130:
        num_col_width -= 30
    elif num_col_width > 90:
        num_col_width -= 22
    elif num_col_width > 80:
        num_col_width -= 16
    elif num_col_width > 70:
        num_col_width -= 13
    elif num_col_width > 60:
        num_col_width -= 10
    elif num_col_width > 50:
        num_col_width -= 6
    elif num_col_width > 40:
        num_col_width -= 4

    return num_col_width

def get_sigpath_sel_trx_group_str(tech_sel_et):
    trx_group = [[], [], [], [], [], [], [], [], [], []]
    selection_path_group_et = tech_sel_et.find('sig_path_selection_group')
    if isinstance(selection_path_group_et, ET.Element):
        sel_path_group_et = selection_path_group_et.find('sig_path_sel_group')
        group_list = sel_path_group_et.findall('group')

        if group_list:
            for group_et in group_list:
                tx_group_list = group_et.findall('tx_operation')
                group_index = 0
                if tx_group_list:
                    for tx_et in tx_group_list:
                        tx_info_s = ''
                        tx_info_s += tx_et.find('band').attrib['name']
                        tx_info_s += '['
                        tx_info_s += tx_et.find('band').find('tx').find('sig_path').text
                        tx_info_s += ']'
                        trx_group[group_index].append(tx_info_s)
                        group_index += 1

                for tx_num in range(group_index, 2):
                    trx_group[tx_num].append('')

                rx_group_list = group_et.findall('rx_operation')
                group_index = 2
                if rx_group_list:
                    for rx_et in rx_group_list:
                        rx_info_s = ''
                        rx_info_s += rx_et.find('band').attrib['name']
                        rx_info_s += '['
                        rx_sig_path_list = rx_et.find('band').find('rx').findall('sig_path')
                        for rx_sig_path_et in rx_sig_path_list:
                            rx_info_s += rx_sig_path_et.text
                            rx_info_s += ' '
                        rx_info_s = rx_info_s[0:len(rx_info_s) - 1]
                        rx_info_s += ']'
                        trx_group[group_index].append(rx_info_s)
                        group_index += 1
                    for rx_num in range(group_index, 10):
                        trx_group[rx_num].append('')
    return trx_group

def get_antpath_restriction_tbl(allow_disallow_et):
    antpath_restriction = [[], []]
    antpath_group_list = allow_disallow_et.findall('group')

    if antpath_group_list:
        for antpath_group_et in antpath_group_list:
            antpath_s = ''
            antpath_a_list = antpath_group_et.find('sig_path_a').findall('sig_path')
            for antpath_a_et in antpath_a_list:
                antpath_s += antpath_a_et.text + ' '
            antpath_s = antpath_s[0:len(antpath_s) - 1]
            antpath_restriction[0].append(antpath_s)

            antpath_s = ''
            antpath_b_list = antpath_group_et.find('sig_path_b').findall('sig_path')
            for antpath_b_et in antpath_b_list:
                antpath_s += antpath_b_et.text + ' '
            antpath_s = antpath_s[0:len(antpath_s) - 1]
            antpath_restriction[1].append(antpath_s)
    return antpath_restriction

def get_antpath_sel_id_str(trx_group_et, tx_rx_s):

    ant_group_str = ''
    id_list = trx_group_et.findall(tx_rx_s)
    ant_group_str += 'BGN: ' + id_list[0].find('bandclass_id').text + ' \n'
    for id_et in id_list:
        ant_group_str += 'ID: ' + id_et.attrib['id'] + ' '
        ant_group_str += 'SPG: ' + id_et.find('sig_path_group_id').text + ' ' + 'ASP: '
        ASP_list = id_et.find('override_antswitch_path').findall('ant_switching_config')
        ASP_str = ''
        for ASP_et in ASP_list:
            ASP_str += ASP_et.find('ant_path').text + ' '
        ASP_str = ASP_str[0:len(ASP_str) - 1]
        ant_group_str += ASP_str + '\n'
    ant_group_str = ant_group_str[0:len(ant_group_str) - 1]

    return ant_group_str

'''
Python etree handle XML file
'''
# ---------------------------------function--------------------------------
# print the element node info for better understandings
# -------------------------------------------------------------------------
def print_tree_element_info(element, namestring):
    print('%s type' % namestring, type(element))
    print('%s tag' % namestring, '=========', element.tag, '=========')
    print('%s length:' % namestring, len(element))
# -------------------------------------------------------------------------

rfc_path = input("Input RFC XML: ")

# file element tree and get root node
tree= ET.ElementTree()
try:
    tree.parse(rfc_path)
    print(rfc_path)
except:
    print('Use default RFC path')
    tree.parse(r'C:\CODE\MPSS.HI.1.0.c8-00198\modem_proc\rf\rfc_himalaya\common\etc\rf_card\rfc_Global_SDRV300_BoardID2_ag.xml')

# initial excel table
wb = XW.Workbook('RFC.xlsx')
ws_cardvariants = wb.add_worksheet('Card Variants')
ws_cardvariants.freeze_panes(2, 0)
ws_phydevice = wb.add_worksheet('Physical Devices')
ws_phydevice.freeze_panes(0, 1)
ws_signalpath = wb.add_worksheet('Signal Path')
ws_signalpath.freeze_panes(1, 1)
ws_antpath = wb.add_worksheet('Ant Switch Path')
ws_antpath.freeze_panes(1, 1)
ws_fbrx = wb.add_worksheet('FBRX Path')
ws_fbrx.freeze_panes(1, 1)


format1 = wb.add_format({'font_size': 10, 'font_name': 'Calibri', 'bold': 1, 'font_color': 'white',
                         'fg_color': 'green',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1}) # 自动换行
format2 = wb.add_format({'font_size': 9, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': '#F2F2F2',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1, 'shrink': 0})
format3 = wb.add_format({'font_size': 9, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': 'white',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})
format4 = wb.add_format({'font_size': 9, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': '#F2F2F2',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 0, 'shrink': 1})
format5 = wb.add_format({'font_size': 9, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': 'white',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 0, 'shrink': 1})
format6 = wb.add_format({'font_size': 11, 'font_name': 'Calibri', 'bold': 1, 'font_color': 'red',
                         'fg_color': 'yellow',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1}) # 自动换行

root=tree.getroot()

# root RFC
print('root type', type(root))
print('root tag','=========',root.tag,'=========')
print('root length:', len(root))
assert isinstance(root, ET.Element)

#================ @ Card Variants ===================
child_antpath = root.find("card_variants")
if isinstance(child_antpath, ET.Element):
    card_properties_et = child_antpath.find('card_properties')
    name_s = ''
    if isinstance(card_properties_et.find('name'), ET.Element):
        name_s = card_properties_et.find('name').text
    hwid_s = ''
    if isinstance(card_properties_et.find('hwid'), ET.Element):
        hwid_s = card_properties_et.find('hwid').text
    swid_s = ''
    if isinstance(card_properties_et.find('fsid'), ET.Element):
        swid_s = card_properties_et.find('fsid').text
    board_id_s = ''
    if isinstance(card_properties_et.find('board_id'), ET.Element):
        board_id_s = card_properties_et.find('board_id').text
    protection_level_s = ''
    if isinstance(card_properties_et.find('protection_level'), ET.Element):
        protection_level_s = card_properties_et.find('protection_level').text
    target_s = ''
    target_list_et = card_properties_et.find('target_list')
    if isinstance(target_list_et, ET.Element):
        for target_et in target_list_et.findall('target'):
            target_s += target_et.text + ','
        target_s = target_s[0:len(target_s)-1]

    Title_s = ['Name', 'HWID', 'SWID', 'Board ID', 'Targets', 'Protection Level']
    values = [name_s, hwid_s, swid_s, board_id_s, target_s, protection_level_s]

    for index_card in range(0, len(Title_s)):
        ws_cardvariants.write(0, index_card, Title_s[index_card], format1)
        ws_cardvariants.write(1, index_card, values[index_card], format3)
        ws_cardvariants.set_column(index_card,index_card, max(calc_col_width_by_str(Title_s[index_card]) + 1, calc_col_width_by_str(values[index_card])))
    ws_cardvariants.set_row(0,20)

#================ @ get ant path ====================
child_antpath = root.find("ant_switch_paths")
assert isinstance(child_antpath, ET.Element)
antpath_to_antnum = {}

Heading_list_ant = ['Path ID', 'Antenna', 'Tuner0', 'Tuner1', 'Tuner2', 'Tuner3', 'XSW0', 'XSW1', 'XSW2',
                    'GRFC XSW0', 'GRFC XSW1', 'GRFC XSW2']
col_width_ant = []

for str_write in Heading_list_ant:
    col_width_ant.append(calc_col_width_by_str(str_write) + 0.5)

ws_antpath.write_row(0, 0, Heading_list_ant, format1)
row_ant = 1

for ant_path_et in child_antpath:
    antpath_id = ant_path_et.attrib['path_id']
    ant_num = ''
    tuner0 = ''
    tuner1 = ''
    tuner2 = ''
    tuner3 = ''
    xsw0 = ''
    xsw1 = ''
    xsw2 = ''
    grfc_xsw0 = ''
    grfc_xsw1 = ''
    grfc_xsw2 = ''

    ant_num_et = ant_path_et.find('antenna')
    if isinstance(ant_num_et, ET.Element):
        ant_num = ant_num_et.text

    module_list_et1 = ant_path_et.find('module_list')
    if isinstance(module_list_et1, ET.Element):
        tuner_list = module_list_et1.findall('tuner')
        if tuner_list:
            tuner_num = len(tuner_list)
            tuner_s_list = []
            for tuner_et in tuner_list:
                tuner_s_list.append(tuner_et.attrib['module_id'])
            tuner0 = tuner_s_list[0]
            if tuner_num > 1:
                tuner1 = tuner_s_list[1]
            if tuner_num > 2:
                tuner2 = tuner_s_list[2]
            if tuner_num > 3:
                tuner3 = tuner_s_list[3]
        xsw_list = module_list_et1.findall('xsw')
        if xsw_list:
            xsw_num = len(xsw_list)
            xsw_s_list = []
            for xsw_et_i in xsw_list:
                xsw_s_i = xsw_et_i.attrib['module_id']
                xsw_s_i += '\n' + xsw_et_i.find('port').text
                xsw_s_list.append(xsw_s_i)
            xsw0 = xsw_s_list[0]
            if xsw_num > 1:
                xsw1 = xsw_s_list[1]
            if xsw_num > 2:
                xsw2 = xsw_s_list[2]
        grfc_xsw_list = module_list_et1.findall('grfc_xsw')
        if grfc_xsw_list:
            grfc_xsw_num = len(grfc_xsw_list)
            grfc_xsw_s_list = []
            for grfc_xsw_i in grfc_xsw_list:
                grfc_xsw_cfg_list = grfc_xsw_i.findall('grfc_config')
                grfc_xsw_i_s = ''
                grfc_xsw_i_s += grfc_xsw_i.attrib['module_id'] + '\n'
                for grfc_xsw_cfg_et in grfc_xsw_cfg_list:
                    grfc_xsw_i_s += grfc_xsw_cfg_et.find('grfc_type').text + '\n'
                    grfc_xsw_i_s += grfc_xsw_cfg_et.find('signal').text + '\n'
                    grfc_xsw_i_s += grfc_xsw_cfg_et.find('enable').text + '\n'
                    grfc_xsw_i_s += grfc_xsw_cfg_et.find('disable').text + '\n'
                grfc_xsw_i_s = grfc_xsw_i_s[0:len(grfc_xsw_i_s)-1]
                grfc_xsw_s_list.append(grfc_xsw_i_s)
            grfc_xsw0 = grfc_xsw_s_list[0]
            if grfc_xsw_num > 1:
                grfc_xsw1 = grfc_xsw_s_list[1]
            if grfc_xsw_num > 2:
                grfc_xsw2 = grfc_xsw_s_list[2]

    antpath_to_antnum[antpath_id] = ant_num
    row_list_ant =[
        antpath_id,
        ant_num,
        tuner0,
        tuner1,
        tuner2,
        tuner3,
        xsw0,
        xsw1,
        xsw2,
        grfc_xsw0,
        grfc_xsw1,
        grfc_xsw2
    ]

    for col_ant_i in range(len(row_list_ant)):
        ws_antpath.write(row_ant, col_ant_i, row_list_ant[col_ant_i], format3)
        col_width_n_ant = calc_col_width_by_str(row_list_ant[col_ant_i])
        if col_width_n_ant > col_width_ant[col_ant_i]:
            col_width_ant[col_ant_i] = col_width_n_ant
    row_ant += 1

    print('antpath_id', antpath_id,
          'ant_num', ant_num,
          'tuner', tuner0, tuner1, tuner2, tuner3,
          'XSW', xsw0, xsw1, xsw2,
          'GRFC XSW', grfc_xsw0, grfc_xsw1, grfc_xsw2)

for coln in range(len(col_width_ant)):
    ws_antpath.set_column(coln, coln, col_width_ant[coln])
ws_antpath.autofilter(0, 0, row_ant - 1, len(col_width_ant))

#================ @get sig_paths node ===================
child=root.find("sig_paths")
print_tree_element_info(child,'child')
assert isinstance(child, ET.Element)

Heading_list = ['Path ID', 'Type\n(rx/tx)', 'PRX?', 'Max\nTX BW', 'PWR\nClass', 'Functionality',
                'Cal Ref\nsigpath', 'Path\noverride\nIndex', 'MCS\n256QAM', 'Disabled', 'FBRX', 'Sigpath\nGroup',
                'Ant SW\npath', 'ASP', 'Split band\nsigpath Map', 'Split Band', 'Tech Bands', 'TRX', 'ELNA0', 'ELNA1', 'PA', 'PAPM', 'PAPM HUB',
                'ASM0', 'ASM1', 'ASM2', 'ASM3', 'ASM4', 'GRFC ASM0', 'GRFC ASM1', 'GRFC ASM2', 'THERM', 'THERM MITIGATION']
col_width_s = []

for str_write in Heading_list:
    col_width_s.append(calc_col_width_by_str(str_write) + 0.5)

# ws_signalpath.set_row(0, 40)
ws_signalpath.write_row(0, 0, Heading_list, format1)
row_n = 1

for sig_path_i in child:
    assert isinstance(sig_path_i, ET.Element)
    path_id = sig_path_i.attrib['path_id']
    path_type = ''
    is_PRX = False
    max_tx_bw = ''
    pwr_class = ''
    functionality = ''
    cal_ref = ''
    path_override_idx = ''
    MCS_256QAM = ''
    Disabled = ''
    fbrx = ''
    SPG = ''
    ant_sw_path = ''
    ASP = ''
    split_band_sig_path_map = ''
    split_band_channel_range = ''
    tech_bands = ''
    TRX = ''
    ELNA0 = ''
    ELNA1 = ''
    PA = ''
    PAPM = ''
    PAPM_HUB = ''
    ASM0 = ''
    ASM1 = ''
    ASM2 = ''
    ASM3 = ''
    ASM4 = ''
    grfc_ASM0 = ''
    grfc_ASM1 = ''
    grfc_ASM2 = ''
    therm = ''
    therm_mitigation = ''

    path_type_et = sig_path_i.find('path_type')
    if isinstance(path_type_et, ET.Element):
        path_type = path_type_et.text

    is_prx_et = sig_path_i.find('sig_path_preferred_prx')
    if isinstance(is_prx_et, ET.Element):
        is_PRX = True

    max_tx_bw_et = sig_path_i.find('max_tx_bw_mhz')
    if isinstance(max_tx_bw_et, ET.Element):
        max_tx_bw = max_tx_bw_et.text

    pwr_class_et = sig_path_i.find('power_class')
    if isinstance(pwr_class_et, ET.Element):
        pwr_class = pwr_class_et.text

    functionality_et = sig_path_i.find('functionality')
    if isinstance(functionality_et, ET.Element):
        functionality = functionality_et.text

    cal_ref_et = sig_path_i.find('cal_reference_sig_path')
    if isinstance(cal_ref_et, ET.Element):
        cal_ref = cal_ref_et.text

    path_override_idx_et = sig_path_i.find('path_override_idx')
    if isinstance(path_override_idx_et, ET.Element):
        path_override_idx = path_override_idx_et.text

    MCS_256QAM_et = sig_path_i.find('mcs_256qam_supported')
    if isinstance(MCS_256QAM_et, ET.Element):
        MCS_256QAM = MCS_256QAM_et.text

    Disabled_et = sig_path_i.find('disabled_on_card_variants')
    if isinstance(Disabled_et, ET.Element):
        for Disabled_et_i in Disabled_et:
            Disabled += Disabled_et_i.text
            if Disabled_et_i != Disabled_et[-1]:
                Disabled += '\n'

    fbrx_et = sig_path_i.find('fbrx_path_assn_properties')
    if isinstance(fbrx_et, ET.Element):
        fbrx_et_child = fbrx_et.find('fbrx_path_assn_config')
        fbrx_et_child_child = fbrx_et_child.find('fbrx_path_assn')
        fbrx = fbrx_et_child_child.text

    SPG_et = sig_path_i.find('sig_path_group_id')
    if isinstance(SPG_et, ET.Element):
        SPG = SPG_et.text

    ant_sw_path_et = sig_path_i.find('ant_switching_properties')
    if isinstance(ant_sw_path_et, ET.Element):
        ant_num = len(ant_sw_path_et)
        ant_i = 0
        for ant_sw_i in ant_sw_path_et:
            ant_i += 1
            ant_sw_path += ant_sw_i[0].text
            ASP += antpath_to_antnum[ant_sw_i[0].text]
            if ant_i < ant_num:
                ant_sw_path += ' '
                ASP += ' '

    band_split_channel_list_et = sig_path_i.find('band_split_channel_list')
    if isinstance(band_split_channel_list_et, ET.Element):
        sigpath_mapping_et = band_split_channel_list_et.findall('split_band_sig_path_mapping')
        if sigpath_mapping_et:
            split_band_sig_path_map = ' '
            for sigpath_mapping_et_i in sigpath_mapping_et:
                split_band_sig_path_map += (sigpath_mapping_et_i[0].text + ' ')
        channel_range_et = band_split_channel_list_et.findall('channel_range')
        if channel_range_et:
            for i in range(len(channel_range_et)):
                start_channel = channel_range_et[i].find('start_channel').text
                end_channel = channel_range_et[i].find('stop_channel').text
                bandwidth = channel_range_et[i].find('bandwidth').text
                split_band_channel_range += (start_channel + ',' + end_channel + ',' + bandwidth)
                if i < len(channel_range_et) - 1:
                    split_band_channel_range += '\n'

    applicable_bands_et = sig_path_i.find('applicable_bands')
    if isinstance(applicable_bands_et, ET.Element):
        for tech_et in applicable_bands_et:
            tech_name = tech_et.attrib['tech_type']
            band_list = tech_et.findall('band')
            for band_i in band_list:
                band_name = band_i.find('band_name').text
                sub_band = band_i.find('sub_band')
                sub_band_s = ''
                if isinstance(sub_band, ET.Element):
                    sub_band_s = sub_band.text
                cal_ant_sw_cnf = band_i.findall('cal_info_per_ant_switching_config')
                cal_flag = {'NO_CAL': 'NC', 'FULL_CAL': 'FC'}
                cal_ant_sw_cnf_s = '['
                for ant_i in range(len(cal_ant_sw_cnf)):
                    cal_ant_sw_cnf_s += cal_flag[cal_ant_sw_cnf[ant_i].text]
                    if ant_i < len(cal_ant_sw_cnf) - 1:
                        cal_ant_sw_cnf_s += ' '

                cal_ant_sw_cnf_s += ']'
                band_s = tech_name + ' ' + band_name + sub_band_s + cal_ant_sw_cnf_s + '\n'
                tech_bands += band_s
        # tech_bands.rstrip()  # remove the end '\n' of the str
        tech_bands = tech_bands[0:len(tech_bands)-1]

    module_list_et = sig_path_i.find('module_list')
    if isinstance(module_list_et, ET.Element):
        TRX_et = module_list_et.find('trx')
        if isinstance(TRX_et, ET.Element):
            TRX += TRX_et.attrib['module_id']
            TRX += '\n'
            port_et = TRX_et.find('port')
            if isinstance(port_et, ET.Element):
                TRX += port_et.text
                TRX += '\n'
            tx_gain_et = TRX_et.find('tx_gain_lineups')
            tx_gain_s = ''
            if isinstance(tx_gain_et, ET.Element):
                gain_lineup_et = tx_gain_et[0]
                for gain_state in gain_lineup_et:
                    tx_gain_s += gain_state.text
                TRX += (tx_gain_s + '\n')
            TRX = TRX[0:len(TRX) - 1]
        ELNA_et_list = module_list_et.findall('elna')
        if ELNA_et_list:
            elna_num = len(ELNA_et_list)
            ELNA_n = []
            for ELNA_et in ELNA_et_list:
                ELNA_s = ''
                elna_port = ELNA_et.find('port')
                ELNA_s += (ELNA_et.attrib['module_id'] + '\n')
                if isinstance(elna_port, ET.Element):
                    ELNA_s += (elna_port.text + '\n')
                rx_gain_lineups = ELNA_et.find('rx_gain_lineups')
                if isinstance(rx_gain_lineups, ET.Element):
                    for gain_lineup in rx_gain_lineups:
                        gain_lineup_tech = gain_lineup.attrib['tech']
                        gain_line_s = ''
                        for gain_state in gain_lineup:
                            gain_line_s += (gain_state.text+ ' ')
                        ELNA_s += (gain_lineup_tech + ' ' + gain_line_s + '\n')
                ELNA_s = ELNA_s[0:len(ELNA_s) - 1]
                ELNA_n.append(ELNA_s)
            ELNA0 = ELNA_n[0]
            if elna_num>1:
                ELNA1 = ELNA_n[1]
        PA_et = module_list_et.find('pa')
        if isinstance(PA_et, ET.Element):
            PA += (PA_et.attrib['module_id'] + '\n')
            PA += PA_et.find('port').text
        PAPM_et = module_list_et.find('papm')
        if isinstance(PAPM_et, ET.Element):
            PAPM += (PAPM_et.attrib['module_id'] + '\n')
            PAPM += PAPM_et.find('port').text
        PAPM_HUB_et = module_list_et.find('papm_hub')
        if isinstance(PAPM_HUB_et, ET.Element):
            PAPM_HUB += (PAPM_HUB_et.attrib['module_id'] + '\n')
            PAPM_HUB += PAPM_HUB_et.find('port').text
        ASM_et_list = module_list_et.findall('asm')
        if ASM_et_list:
            ASM_num = len(ASM_et_list)
            ASM_n = []
            for ASM_et in ASM_et_list:
                ASM_s = ''
                ASM_s += (ASM_et.attrib['module_id'] + '\n')
                ASM_s += ASM_et.find('port').text
                ASM_n.append(ASM_s)
            ASM0 = ASM_n[0]
            if ASM_num > 1:
                ASM1 = ASM_n[1]
            if ASM_num > 2:
                ASM2 = ASM_n[2]
            if ASM_num > 3:
                ASM3 = ASM_n[3]
            if ASM_num > 4:
                ASM4 = ASM_n[4]
        grfc_ASM_et_list = module_list_et.findall('grfc_asm')
        if grfc_ASM_et_list:
            grfc_ASM_num = len(grfc_ASM_et_list)
            grfc_ASM_n = []
            for grfc_ASM_et in grfc_ASM_et_list:
                grfc_ASM_s = ''
                grfc_ASM_s += (grfc_ASM_et.attrib['module_id'] + '\n')
                grfc_config_list = grfc_ASM_et.findall('grfc_config')
                for grfc_config_et in grfc_config_list:
                    grfc_ASM_s += grfc_config_et.find('signal').text + '\n'
                    grfc_ASM_s += grfc_config_et.find('grfc_type').text + '\n'
                    grfc_ASM_s += grfc_config_et.find('enable').text + '\n'
                    grfc_ASM_s += grfc_config_et.find('disable').text + '\n'
                grfc_ASM_s = grfc_ASM_s[0:len(grfc_ASM_s)-1]
                grfc_ASM_n.append(grfc_ASM_s)
            grfc_ASM0 = grfc_ASM_n[0]
            if grfc_ASM_num > 1:
                grfc_ASM1 = grfc_ASM_n[1]
            if grfc_ASM_num > 2:
                grfc_ASM2 = grfc_ASM_n[2]
        # GRFC PA
        # GRFC ELNA
        therm_et = module_list_et.find('therm')
        if isinstance(therm_et, ET.Element):
            therm = therm_et.attrib['module_id']
        therm_mitigation_et = module_list_et.find('therm_mitigation')
        if isinstance(therm_mitigation_et, ET.Element):
            therm_mitigation = therm_mitigation_et.attrib['module_id']

    row_list_sigpath = [path_id,
                        path_type,
                        str(is_PRX),
                        max_tx_bw,
                        pwr_class,
                        functionality,
                        cal_ref,
                        path_override_idx,
                        MCS_256QAM,
                        Disabled,
                        fbrx,
                        SPG,
                        ant_sw_path,
                        ASP,
                        split_band_sig_path_map,
                        split_band_channel_range,
                        tech_bands,
                        TRX,
                        ELNA0,
                        ELNA1,
                        PA,
                        PAPM,
                        PAPM_HUB,
                        ASM0,
                        ASM1,
                        ASM2,
                        ASM3,
                        ASM4,
                        grfc_ASM0,
                        grfc_ASM1,
                        grfc_ASM2,
                        therm,
                        therm_mitigation]

    for col_n in range(len(row_list_sigpath)):
        if col_n == 9: # Disabled 这一栏不重要 缩小显示
            if path_type == 'tx':
                format_col = format4
            else:
                format_col = format5
            ws_signalpath.write(row_n, col_n, row_list_sigpath[col_n], format_col)
        else:
            if path_type == 'tx':
                format_col = format2
            else:
                format_col = format3
            ws_signalpath.write(row_n, col_n, row_list_sigpath[col_n], format_col)
            col_width_n = calc_col_width_by_str(row_list_sigpath[col_n])
            if col_width_n > col_width_s[col_n]:
                col_width_s[col_n] = col_width_n

    row_n += 1

    '''
    ws_signalpath.write(row_n, 0, path_id, format2)
    ws_signalpath.write(row_n, 1, path_type, format2)
    ws_signalpath.write(row_n, 2, is_PRX, format2)
    ws_signalpath.write(row_n, 3, max_tx_bw, format2)
    ws_signalpath.write(row_n, 4, pwr_class, format2)
    ws_signalpath.write(row_n, 5, functionality, format2)
    ws_signalpath.write(row_n, 6, cal_ref, format2)
    ws_signalpath.write(row_n, 7, path_override_idx, format2)
    ws_signalpath.write(row_n, 8, MCS_256QAM, format2)
    ws_signalpath.write(row_n, 9, Disabled, format2)
    ws_signalpath.write(row_n, 10, fbrx, format2)
    ws_signalpath.write(row_n, 11, SPG, format2)
    ws_signalpath.write(row_n, 12, ant_sw_path, format2)
    ws_signalpath.write(row_n, 13, split_band_sig_path_map, format2)
    ws_signalpath.write(row_n, 14, split_band_channel_range, format2)
    ws_signalpath.write(row_n, 15, tech_bands, format2)
    ws_signalpath.write(row_n, 16, TRX, format2)
    ws_signalpath.write(row_n, 17, ELNA0, format2)
    ws_signalpath.write(row_n, 18, ELNA1, format2)
    ws_signalpath.write(row_n, 19, PA, format2)
    ws_signalpath.write(row_n, 20, PAPM, format2)
    ws_signalpath.write(row_n, 21, PAPM_HUB, format2)
    ws_signalpath.write(row_n, 22, ASM0, format2)
    ws_signalpath.write(row_n, 23, ASM1, format2)
    ws_signalpath.write(row_n, 24, ASM2, format2)
    ws_signalpath.write(row_n, 25, ASM3, format2)
    ws_signalpath.write(row_n, 26, ASM4, format2)
    ws_signalpath.write(row_n, 27, grfc_ASM0, format2)
    ws_signalpath.write(row_n, 28, grfc_ASM1, format2)
    ws_signalpath.write(row_n, 29, grfc_ASM2, format2)
    ws_signalpath.write(row_n, 30, therm, format2)
    ws_signalpath.write(row_n, 31, therm_mitigation, format2)
    '''


    print(path_id, path_type, is_PRX, max_tx_bw, pwr_class, functionality, cal_ref, path_override_idx, MCS_256QAM, Disabled, fbrx, SPG, '-', ant_sw_path,
          '--', split_band_sig_path_map,
          '--', split_band_channel_range,
          'tech_bands', tech_bands,
          'TRX:',TRX,
          'ELNA0', ELNA0,
          'ELNA1', ELNA1,
          'PA', PA,
          'PAPM', PAPM,
          'PAPM_HUB', PAPM_HUB, 'ASM0',ASM0, 'ASM1',ASM1,'ASM2',ASM2,'ASM3',ASM3,'ASM4',ASM4,
          'grfcASM0',grfc_ASM0, 'grfcASM1',grfc_ASM1,'grfcASM2',grfc_ASM2, 'therm', therm, 'therm_mitigation', therm_mitigation)
    print('---------------------------------')

for coln in range(len(col_width_s)):
    ws_signalpath.set_column(coln, coln, col_width_s[coln])
ws_signalpath.autofilter(0, 0, row_n, len(col_width_s))

#================ @ get FBRX path ====================
child_fbrx_path = root.find("fbrx_paths")
assert isinstance(child_fbrx_path, ET.Element)
Heading_list_fbrx = ['Path\nID', 'TRX', 'ASM', 'ASM1', 'GRFC ASM', 'GRFC ASM1', 'GRFC\nCoupler', 'Coupler0', 'Coupler1',
                    'Coupler2', 'Coupler3', 'Coupler4']
col_width_fbrx = []

for str_write in Heading_list_fbrx:
    col_width_fbrx.append(calc_col_width_by_str(str_write) + 0.5)

ws_fbrx.write_row(0, 0, Heading_list_fbrx, format1)
row_fbrx = 1

for fbrx_path_i in child_fbrx_path:
    assert isinstance(fbrx_path_i, ET.Element)
    fbrx_pathid = fbrx_path_i.attrib['path_id']
    fbrx_trx = ''
    fbrx_asm = ''
    fbrx_asm1 = ''
    grfc_asm = ''
    grfc_asm1 = ''
    grfc_coupler = ''
    coupler0 = ''
    coupler1 = ''
    coupler2 = ''
    coupler3 = ''
    coupler4 = ''

    module_list_fbrx_et = fbrx_path_i.find('module_list')

    trx_fbrx_et = module_list_fbrx_et.find('trx')
    fbrx_trx += trx_fbrx_et.attrib['module_id'] + '\n'
    fbrx_trx += trx_fbrx_et.find('port').text

    asm_fbrx_list = module_list_fbrx_et.findall('asm')
    if asm_fbrx_list:
        fbrx_asm_num = len(asm_fbrx_list)
        fbrx_asm_s_list = []
        for asm_fbrx_i in asm_fbrx_list:
            fbrx_asm_s = asm_fbrx_i.attrib['module_id'] + '\n'
            fbrx_asm_s += asm_fbrx_i.find('port').text
            fbrx_asm_s_list.append(fbrx_asm_s)
        fbrx_asm = fbrx_asm_s_list[0]
        if fbrx_asm_num > 1:
            fbrx_asm1 = fbrx_asm_s_list[1]

    grfc_asm_fbrx_list = module_list_fbrx_et.findall('grfc_asm')
    if grfc_asm_fbrx_list:
        grfc_asm_num_fbrx = len(grfc_asm_fbrx_list)
        grfc_asm_fbrx_s_list = []
        for grfc_asm_fbrx_i in grfc_asm_fbrx_list:
            grfc_asm_fbrx_s = grfc_asm_fbrx_i.attrib['module_id'] + '\n'
            grfc_config_fbrx = grfc_asm_fbrx_i.find('grfc_config')
            grfc_asm_fbrx_s += grfc_config_fbrx.find('grfc_type').text + '\n'
            grfc_asm_fbrx_s += grfc_config_fbrx.find('signal').text + '\n'
            grfc_asm_fbrx_s += grfc_config_fbrx.find('enable').text + '\n'
            grfc_asm_fbrx_s += grfc_config_fbrx.find('disable').text
            grfc_asm_fbrx_s_list.append(grfc_asm_fbrx_s)
        grfc_asm = grfc_asm_fbrx_s_list[0]
        if grfc_asm_num_fbrx > 1:
            grfc_asm1 = grfc_asm_fbrx_s_list[1]

    # grfc_coupler

    coupler_list = module_list_fbrx_et.findall('coupler')
    if coupler_list:
        coupler_num = len(coupler_list)
        coupler_s_list = []
        for coupler_i in coupler_list:
            coupler_s = coupler_i.attrib['module_id'] + '\n'
            coupler_s += coupler_i.find('port').text + '\n'
            pos = coupler_i.find('position')
            if isinstance(pos, ET.Element):
                coupler_s += pos.text + '\n'
            atten_fwd = coupler_i.find('atten_fwd')
            if isinstance(atten_fwd, ET.Element):
                coupler_s += atten_fwd.text + '\n'
            atten_rev = coupler_i.find('atten_rev')
            if isinstance(atten_rev, ET.Element):
                coupler_s += atten_rev.text + '\n'
            coupler_s = coupler_s[0:len(coupler_s)-1]
            coupler_s_list.append(coupler_s)
        coupler0 = coupler_s_list[0]
        if coupler_num > 1:
            coupler1 = coupler_s_list[1]
        if coupler_num > 2:
            coupler2 = coupler_s_list[2]
        if coupler_num > 3:
            coupler3 = coupler_s_list[3]
        if coupler_num > 4:
            coupler4 = coupler_s_list[4]

    row_list_fbrx = [
        fbrx_pathid,
        fbrx_trx,
        fbrx_asm,
        fbrx_asm1,
        grfc_asm,
        grfc_asm1,
        grfc_coupler,
        coupler0,
        coupler1,
        coupler2,
        coupler3,
        coupler4,
    ]

    for col_fbrx_i in range(len(row_list_fbrx)):
        ws_fbrx.write(row_fbrx, col_fbrx_i, row_list_fbrx[col_fbrx_i], format3)
        col_width_n_fbrx = calc_col_width_by_str(row_list_fbrx[col_fbrx_i])
        if col_width_n_fbrx > col_width_fbrx[col_fbrx_i]:
            col_width_fbrx[col_fbrx_i] = col_width_n_fbrx
    row_fbrx += 1

    print('fbrx_pathid', fbrx_pathid, 'fbrx_trx',fbrx_trx, 'fbrx_asm', fbrx_asm, fbrx_asm1, 'grfc_asm', grfc_asm, grfc_asm1,
          'coupler', coupler0, coupler1, coupler2, coupler3, coupler4)

for coln in range(len(col_width_fbrx)):
    ws_fbrx.set_column(coln, coln, col_width_fbrx[coln])
ws_fbrx.autofilter(0, 0, row_fbrx - 1, len(col_width_fbrx))


#================ @ get physical device list ====================
child_physical_device = root.find("phy_device_list")

index = 0
GRFC_device_list = []
RFFE_device_list = []
QLINK_list = []
Alt_rffe_device_list = []
for device_et in child_physical_device:
    device_tag = device_et.tag
    device_str_list = []
    device_str_list.append(index)
    if device_tag == 'device':
        type = device_et.attrib['type']
        if re.search('GRFC', type):
            device_str_list.append(type)
            grfc_et = device_et.find('grfc')
            if isinstance(grfc_et, ET.Element):
                device_str_list.append(grfc_et.find('comm_master').text)
            grfc_module_et = device_et.find('module_list')
            Mod0 = ''
            Mod1 = ''
            Mod2 = ''
            Mod_list = []
            for module_et in grfc_module_et:
                Mod_list.append(module_et.attrib['id']+'\n'+module_et.find('type').text)
            grfc_Mod_num = len(Mod_list)
            if grfc_Mod_num > 0:
                Mod0 = Mod_list[0]
            if grfc_Mod_num > 1:
                Mod1 = Mod_list[1]
            if grfc_Mod_num > 2:
                Mod2 = Mod_list[2]
            device_str_list.append(Mod0)
            device_str_list.append(Mod1)
            device_str_list.append(Mod2)
            GRFC_device_list.append(device_str_list)
        elif re.search('SDR', type):
            device_str_list.append(type)
            qlink_et = device_et.find('qlink')
            if isinstance(qlink_et, ET.Element):
                device_str_list.append(qlink_et.find('channel').text)
            qlink_module_list = device_et.find('module_list')
            mod0 = ''
            mod1 = ''
            mod_list = []
            for mod_et in qlink_module_list:
                mod_list.append(mod_et.attrib['id']+'\n'+mod_et.find('type').text)
            mod0 = mod_list[0]
            if len(mod_list) > 1:
                mod1 = mod_list[1]
            device_str_list.append(mod0)
            device_str_list.append(mod1)
            QLINK_list.append(device_str_list)

        else:
            device_str_list.append(type)
            rffe = device_et.find('rffe')
            if isinstance(rffe, ET.Element):
                device_str_list.append(rffe.find('protocol_version').text)
                device_str_list.append(rffe.find('comm_master').text)
                device_str_list.append(rffe.find('channel').text)
                device_str_list.append(rffe.find('manufacturer_id').text)
                device_str_list.append(rffe.find('product_id').text)
                device_str_list.append(rffe.find('product_rev').text)
                device_str_list.append(rffe.find('default_usid').text)
                device_str_list.append(rffe.find('assigned_usid').text)
            module_et_list = device_et.find('module_list')
            mod_list = ['','','','','','','']
            mod_list_s = []
            for module_et in module_et_list:
                id_s = module_et.attrib['id']
                specifier = ''
                specifier_et = module_et.find('specifier')
                if isinstance(specifier_et, ET.Element):
                    specifier = specifier_et.text
                type_mod = ''
                type_mod_et = module_et.find('type')
                if isinstance(type_mod_et, ET.Element):
                    type_mod = type_mod_et.text
                mod_list_s.append(id_s+'\n'+specifier+'\n'+type_mod)
            rffe_mod_num = len(mod_list_s)
            for i_num in range(len(mod_list_s)):
                if i_num < len(mod_list):
                    mod_list[i_num] = mod_list_s[i_num]
            for mod_i in mod_list:
                device_str_list.append(mod_i)

            RFFE_device_list.append(device_str_list)
    elif device_tag == 'alternate_devices':
        device_et_primary = device_et[0]
        device_str_list.append(device_et_primary.attrib['type'])
        rffe = device_et_primary.find('rffe')
        if isinstance(rffe, ET.Element):
            device_str_list.append(rffe.find('protocol_version').text)
            device_str_list.append(rffe.find('comm_master').text)
            device_str_list.append(rffe.find('channel').text)
            device_str_list.append(rffe.find('manufacturer_id').text)
            device_str_list.append(rffe.find('product_id').text)
            device_str_list.append(rffe.find('product_rev').text)
            device_str_list.append(rffe.find('default_usid').text)
            device_str_list.append(rffe.find('assigned_usid').text)
        module_et_list = device_et_primary.find('module_list')
        mod_list = ['', '', '', '', '', '', '']
        mod_list_s = []
        for module_et in module_et_list:
            id_s = module_et.attrib['id']
            specifier = ''
            specifier_et = module_et.find('specifier')
            if isinstance(specifier_et, ET.Element):
                specifier = specifier_et.text
            type_mod = ''
            type_mod_et = module_et.find('type')
            if isinstance(type_mod_et, ET.Element):
                type_mod = type_mod_et.text
            mod_list_s.append(id_s + '\n' + specifier + '\n' + type_mod)
        rffe_mod_num = len(mod_list_s)
        for i_num in range(len(mod_list_s)):
            if i_num < len(mod_list):
                mod_list[i_num] = mod_list_s[i_num]
        for mod_i in mod_list:
            device_str_list.append(mod_i)
        RFFE_device_list.append(device_str_list)

        device_et_alt = device_et[1]
        device_str_list_alt = []
        device_str_list_alt.append(index)
        device_str_list_alt.append(device_et_alt.attrib['type'])
        rffe_alt = device_et_alt.find('rffe')
        device_str_list_alt.append(rffe_alt.find('manufacturer_id').text)
        device_str_list_alt.append(rffe_alt.find('product_id').text)
        device_str_list_alt.append(rffe_alt.find('product_rev').text)

        Alt_rffe_device_list.append(device_str_list_alt)
        print(device_str_list_alt)

    index += 1

    print(device_str_list)

ws_phydevice.write(0, 0, 'RFFE DEVICES', format6)
ws_phydevice.set_row(0, 25)
FirstCol_str = ['Index', 'Physical Device', 'Proc Ver', 'Comm Master', 'Chan', 'MID', 'PID', 'REV', 'Default USID', 'Assigned USID',
                'Logical Device0\nSpecifier\nType', 'Logical Device1\nSpecifier\nType', 'Logical Device2\nSpecifier\nType', 'Logical Device3\nSpecifier\nType',
                'Logical Device4\nSpecifier\nType', 'Logical Device5\nSpecifier\nType', 'Logical Device6\nSpecifier\nType']
ws_phydevice.write_column(1, 0, FirstCol_str, format1)
col_width_list = []
coln_width_record = []
for col_s in FirstCol_str:
    col_width_list.append(calc_col_width_by_str(col_s) + 0.5)
ws_phydevice.set_column(0,0, max(col_width_list))
coln_width_record.append( max(col_width_list))
col_idx = 1

for rffe_device_col in RFFE_device_list:
    if col_idx % 2:
        ws_phydevice.write_column(1, col_idx, rffe_device_col, format3)
    else:
        ws_phydevice.write_column(1, col_idx, rffe_device_col, format2)
    col_width_list = []
    for col_s in rffe_device_col:
        col_width_list.append(calc_col_width_by_str(col_s) + 0.8)
    ws_phydevice.set_column(col_idx, col_idx, max(col_width_list))
    coln_width_record.append(max(col_width_list))
    col_idx += 1

row_idx = len(FirstCol_str) + 1
if Alt_rffe_device_list:
    ws_phydevice.write(row_idx, 0, 'ALT RFFE\nDEVICES', format6)
    ws_phydevice.set_row(row_idx, 30)
    row_idx += 1
    FirstCol_str_ALT = ['Index', 'Type', 'MID', 'PID', 'REV']
    ws_phydevice.write_column(row_idx, 0, FirstCol_str_ALT, format1)
    col_idx = 1

    for alt_device_col in Alt_rffe_device_list:
        if col_idx % 2:
            ws_phydevice.write_column(row_idx, col_idx, alt_device_col, format3)
        else:
            ws_phydevice.write_column(row_idx, col_idx, alt_device_col, format2)

        col_width_list = []
        for col_s in alt_device_col:
            col_width_list.append(calc_col_width_by_str(col_s) + 0.8)
        if max(col_width_list) > coln_width_record[col_idx]:
            ws_phydevice.set_column(col_idx, col_idx, max(col_width_list))
            coln_width_record[col_idx] = max(col_width_list)
        col_idx += 1
    row_idx += len(FirstCol_str_ALT)

if GRFC_device_list:
    ws_phydevice.write(row_idx, 0, 'GRFC', format6)
    ws_phydevice.set_row(row_idx, 25)
    row_idx += 1
    FirstCol_str_GRFC = ['Index', 'Type', 'Comm Master', 'Mod0\nType', 'Mod1\nType','Mod2\nType']
    ws_phydevice.write_column(row_idx, 0, FirstCol_str_GRFC, format1)
    col_idx = 1

    for GRFC_device_col in GRFC_device_list:
        if col_idx % 2:
            ws_phydevice.write_column(row_idx, col_idx, GRFC_device_col, format3)
        else:
            ws_phydevice.write_column(row_idx, col_idx, GRFC_device_col, format2)

        col_width_list = []
        for col_s in GRFC_device_col:
            col_width_list.append(calc_col_width_by_str(col_s) + 0.8)

        if col_idx > len(coln_width_record):
            ws_phydevice.set_column(col_idx, col_idx, max(col_width_list))
            coln_width_record.append(max(col_width_list))
        else:
            if max(col_width_list) > coln_width_record[col_idx]:
                ws_phydevice.set_column(col_idx, col_idx, max(col_width_list))
                coln_width_record[col_idx] = max(col_width_list)
        col_idx += 1
    row_idx += len(FirstCol_str_GRFC)

ws_phydevice.write(row_idx, 0, 'QLINK', format6)
ws_phydevice.set_row(row_idx, 25)
row_idx += 1
FirstCol_str_QLINK = ['Index', 'Type', 'QLINK ID', 'Mod0\nType', 'Mod1\nType']
ws_phydevice.write_column(row_idx, 0, FirstCol_str_QLINK, format1)
col_idx = 1

for QLINK_device_col in QLINK_list:
    if col_idx % 2:
        ws_phydevice.write_column(row_idx, col_idx, QLINK_device_col, format3)
    else:
        ws_phydevice.write_column(row_idx, col_idx, QLINK_device_col, format2)

    col_width_list = []
    for col_s in QLINK_device_col:
        col_width_list.append(calc_col_width_by_str(col_s) + 0.8)
    if max(col_width_list) > coln_width_record[col_idx]:
        ws_phydevice.set_column(col_idx, col_idx, max(col_width_list))
    col_idx += 1

#================ @ get Sub band info list ====================
child_subband = root.find("rfc_sub_band_list")
if isinstance(child_subband, ET.Element):
    ws_subband = wb.add_worksheet('Sub Band Info List')
    ws_subband.freeze_panes(1, 0)
    title_line = ['Band ID', 'Split Type', 'Start Freq(KHz)', 'Stop Freq(KHz)']
    freq_group_s = [[], [], [], []]
    freq_list = child_subband.findall('frequency_range')
    for freq_et in freq_list:
        freq_group_s[0].append(freq_et.find('sub_band_id').text)
        freq_group_s[1].append(freq_et.find('split_type').text)
        freq_group_s[2].append(freq_et.find('start_freq_khz').text)
        freq_group_s[3].append(freq_et.find('stop_freq_khz').text)

    subband_row = 0
    subband_col = 0
    ws_subband.set_row(0, 25)
    for subband_col in range(0, len(title_line)):
        subband_row = 0
        ws_subband.write(subband_row, subband_col, title_line[subband_col], format1)
        col_width = calc_col_width_by_str(title_line[subband_col]) + 0.8
        subband_row += 1
        for freq_str in freq_group_s[subband_col]:
            ws_subband.write(subband_row, subband_col, freq_str, format3)
            if calc_col_width_by_str(freq_str) > col_width:
                col_width = calc_col_width_by_str(freq_str)
            subband_row += 1
        ws_subband.set_column(subband_col,subband_col,col_width)

#================ @ get GPIO list ====================
child_gpiolist = root.find("gpio_list_v2")

if isinstance(child_gpiolist, ET.Element):
    ws_gpiolist = wb.add_worksheet('SDM GPIO-RFFE Signals')
    ws_gpiolist.freeze_panes(2, 0)
    ws_gpiolist.set_row(0,25)
    ws_gpiolist_row = 0
    ws_gpiolist_col = 0

    rffe_signals_et = child_gpiolist.find('rffe_signals')
    if isinstance(rffe_signals_et, ET.Element):
        rffe_signal_headline = ['Num', 'Speed', 'Clock', 'PULL', 'Strength or Load', 'Data', 'PULL', 'Strength or Load']
        rffe_signal_group_s = [[], [], [], [], [], [], [], []]
        rffe_signal_list = rffe_signals_et.findall('rffe_signal')
        for rffe_et in rffe_signal_list:
            rffe_signal_group_s[0].append(rffe_et.attrib['num'])
            rffe_signal_group_s[1].append(rffe_et.find('speed').text)
            gpio_list = rffe_et.findall('gpio')
            rffe_signal_group_s[2].append(gpio_list[0].attrib['name'])
            rffe_signal_group_s[3].append(gpio_list[0].find('gpio_pull').text)
            rffe_signal_group_s[4].append(gpio_list[0].find('drv_strength').text)
            rffe_signal_group_s[5].append(gpio_list[1].attrib['name'])
            rffe_signal_group_s[6].append(gpio_list[1].find('gpio_pull').text)
            rffe_signal_group_s[7].append(gpio_list[1].find('drv_strength').text)
        ws_gpiolist.merge_range(0, ws_gpiolist_col, 0, ws_gpiolist_col + 7,
                                         'SDM/SDX Driven RFFE Signals(V2)',
                                         format1)
        for tech_col in range(0,8):
            ws_gpiolist_row = 1
            ws_gpiolist.write(ws_gpiolist_row, tech_col + ws_gpiolist_col, rffe_signal_headline[tech_col], format2)
            ws_gpiolist_row += 1
            col_width = calc_col_width_by_str(rffe_signal_headline[tech_col]) + 0.8
            for rffe_str in rffe_signal_group_s[tech_col]:
                ws_gpiolist.write(ws_gpiolist_row, tech_col + ws_gpiolist_col, rffe_str, format3)
                if (calc_col_width_by_str(rffe_str) + 0.8) > col_width:
                    col_width = (calc_col_width_by_str(rffe_str) + 0.8)
                ws_gpiolist_row += 1
            ws_gpiolist.set_column(tech_col + ws_gpiolist_col, tech_col + ws_gpiolist_col, col_width)

        ws_gpiolist_col += 8
        ws_gpiolist.set_column(ws_gpiolist_col, ws_gpiolist_col, 3)
        ws_gpiolist_col += 1



    other_signals_et = child_gpiolist.find('other_signals')
    if isinstance(other_signals_et, ET.Element):
        other_signals_headline = ['Name', 'PULL', 'Strength']
        other_signals_groups_s = [[], [] ,[]]
        gpio_list = other_signals_et.findall('gpio')
        for gpio_et in gpio_list:
            other_signals_groups_s[0].append(gpio_et.attrib['name'])
            if isinstance(gpio_et.find('gpio_pull'), ET.Element):
                other_signals_groups_s[1].append(gpio_et.find('gpio_pull').text)
            else:
                other_signals_groups_s[1].append('')

            if isinstance(gpio_et.find('drv_strength'), ET.Element):
                other_signals_groups_s[2].append(gpio_et.find('drv_strength').text)
            else:
                other_signals_groups_s[2].append('')
        ws_gpiolist.merge_range(0, ws_gpiolist_col, 0, ws_gpiolist_col + 2,
                                'Other Signals Group',
                                format1)
        for tech_col in range(0,3):
            ws_gpiolist_row = 1
            ws_gpiolist.write(ws_gpiolist_row, tech_col + ws_gpiolist_col, other_signals_headline[tech_col], format2)
            ws_gpiolist_row += 1
            col_width = calc_col_width_by_str(other_signals_headline[tech_col]) + 0.8
            for rffe_str in other_signals_groups_s[tech_col]:
                ws_gpiolist.write(ws_gpiolist_row, tech_col + ws_gpiolist_col, rffe_str, format3)
                if 0.8 + calc_col_width_by_str(rffe_str) > col_width:
                    col_width = calc_col_width_by_str(rffe_str) + 0.8
                ws_gpiolist_row += 1
            ws_gpiolist.set_column(tech_col + ws_gpiolist_col, tech_col + ws_gpiolist_col, col_width)

#================ @ get SDR RFFE/GRFC list ====================
child_sdrrffe = root.find("sdr_gpio_list_v2")

if isinstance(child_sdrrffe, ET.Element):
    ws_sdrrffe = wb.add_worksheet('SDR RFFE GRFC Signals')
    ws_sdrrffe.freeze_panes(2, 0)
    ws_sdrrffe.set_row(0, 25)
    ws_sdrrffe_row = 0
    ws_sdrrffe_col = 0

    rffe_signals_et = child_sdrrffe.find('sdr_rffe_signals')
    if isinstance(rffe_signals_et, ET.Element):
        rffe_signal_headline = ['Num', 'Speed', 'Comm Master', 'Clock', 'Clock Load', 'Data', 'Data Load']
        rffe_signal_group_s = [[], [], [], [], [], [], []]
        rffe_signal_list = rffe_signals_et.findall('sdr_rffe_signal')
        for rffe_et in rffe_signal_list:
            rffe_signal_group_s[0].append(rffe_et.attrib['num'])
            rffe_signal_group_s[1].append(rffe_et.find('speed').text)
            rffe_signal_group_s[2].append(rffe_et.find('comm_master').text)
            gpio_list = rffe_et.findall('gpio')
            rffe_signal_group_s[3].append(gpio_list[0].attrib['name'])
            rffe_signal_group_s[4].append(gpio_list[0].find('load').text)
            rffe_signal_group_s[5].append(gpio_list[1].attrib['name'])
            rffe_signal_group_s[6].append(gpio_list[1].find('load').text)
        ws_sdrrffe.merge_range(0, ws_sdrrffe_col, 0, ws_sdrrffe_col + 6, 'SDR RFFE Signals', format1)
        for tech_col in range(0, 7):
            ws_sdrrffe_row = 1
            ws_sdrrffe.write(ws_sdrrffe_row, tech_col + ws_sdrrffe_col, rffe_signal_headline[tech_col], format2)
            ws_sdrrffe_row += 1
            col_width = calc_col_width_by_str(rffe_signal_headline[tech_col]) + 0.8
            for rffe_str in rffe_signal_group_s[tech_col]:
                ws_sdrrffe.write(ws_sdrrffe_row, tech_col + ws_sdrrffe_col, rffe_str, format3)
                if (calc_col_width_by_str(rffe_str) + 0.8) > col_width:
                    col_width = (calc_col_width_by_str(rffe_str) + 0.8)
                ws_sdrrffe_row += 1
            ws_sdrrffe.set_column(tech_col + ws_sdrrffe_col, tech_col + ws_sdrrffe_col, col_width)
        ws_sdrrffe_col += 7
        ws_sdrrffe.set_column(ws_sdrrffe_col, ws_sdrrffe_col, 3)
        ws_sdrrffe_col += 1


    grfc_signals_et = child_sdrrffe.find('sdr_grfc_signals')
    if isinstance(grfc_signals_et, ET.Element):
        grfc_signal_headline = ['Num', 'Comm Master', 'Name', 'Load', 'Pull', 'Common Init']
        grfc_signal_group_s = [[], [], [], [], [], []]
        grfc_signal_list = grfc_signals_et.findall('sdr_grfc_signal')
        for grfc_et in grfc_signal_list:
            grfc_signal_group_s[0].append(grfc_et.attrib['num'])
            grfc_signal_group_s[1].append(grfc_et.find('comm_master').text)
            gpio_et = grfc_et.find('gpio')
            grfc_signal_group_s[2].append(gpio_et.attrib['name'])
            grfc_signal_group_s[3].append(gpio_et.find('load').text)
            grfc_signal_group_s[4].append(gpio_et.find('gpio_pull').text)
            if isinstance(gpio_et.find('gpio_pull'), ET.Element):
                grfc_signal_group_s[5].append(gpio_et.find('common_init').text)
            else:
                grfc_signal_group_s[5].append('')
        ws_sdrrffe.merge_range(0, ws_sdrrffe_col, 0, ws_sdrrffe_col + 5, 'SDR GRFC Signals', format1)
        for tech_col in range(0, 6):
            ws_sdrrffe_row = 1
            ws_sdrrffe.write(ws_sdrrffe_row, tech_col + ws_sdrrffe_col, grfc_signal_headline[tech_col], format2)
            ws_sdrrffe_row += 1
            col_width = calc_col_width_by_str(grfc_signal_headline[tech_col]) + 0.8
            for rffe_str in grfc_signal_group_s[tech_col]:
                ws_sdrrffe.write(ws_sdrrffe_row, tech_col + ws_sdrrffe_col, rffe_str, format3)
                if (calc_col_width_by_str(rffe_str) + 0.8) > col_width:
                    col_width = (calc_col_width_by_str(rffe_str) + 0.8)
                ws_sdrrffe_row += 1
            ws_sdrrffe.set_column(tech_col + ws_sdrrffe_col, tech_col + ws_sdrrffe_col, col_width)
        ws_sdrrffe_col += 6
        ws_sdrrffe.set_column(ws_sdrrffe_col, ws_sdrrffe_col, 3)
        ws_sdrrffe_col += 1

    blanking_signals_et = child_sdrrffe.find('blanking_grfc_signals')
    if isinstance(blanking_signals_et, ET.Element):
        blanking_signal_headline = ['Num', 'Comm Master', 'Signal Name', 'Signal Type', 'Enable', 'Disable', 'TX PWR Th(dB10)', 'Band List']
        blanking_signal_group_s = [[], [], [], [], [], [], [], []]
        blanking_signal_list = blanking_signals_et.findall('blanking_grfc_signal')
        for blanking_et in blanking_signal_list:
            blanking_signal_group_s[0].append(blanking_et.attrib['num'])
            blanking_signal_group_s[1].append(blanking_et.find('comm_master').text)
            gpio_et = blanking_et.find('signal')
            blanking_signal_group_s[2].append(gpio_et.attrib['name'])
            blanking_signal_group_s[3].append(gpio_et.find('signal_type').text)
            blanking_signal_group_s[4].append(gpio_et.find('enable').text)
            blanking_signal_group_s[5].append(gpio_et.find('disable').text)
            blanking_signal_group_s[6].append(gpio_et.find('tx_pwr_th').text)
            band_list = gpio_et.find('band_list').findall('band')
            band_list_s = ''
            for band_et in band_list:
                band_list_s += (band_et.text + ' ')
            blanking_signal_group_s[7].append(band_list_s[0:len(band_list_s)-1])
        ws_sdrrffe.merge_range(0, ws_sdrrffe_col, 0, ws_sdrrffe_col + 7, 'SDR Blanking GRFC Signals', format1)
        for tech_col in range(0, 8):
            ws_sdrrffe_row = 1
            ws_sdrrffe.write(ws_sdrrffe_row, tech_col + ws_sdrrffe_col, blanking_signal_headline[tech_col], format2)
            ws_sdrrffe_row += 1
            col_width = calc_col_width_by_str(blanking_signal_headline[tech_col]) + 0.8
            for rffe_str in blanking_signal_group_s[tech_col]:
                ws_sdrrffe.write(ws_sdrrffe_row, tech_col + ws_sdrrffe_col, rffe_str, format3)
                if (calc_col_width_by_str(rffe_str) + 0.8) > col_width:
                    col_width = (calc_col_width_by_str(rffe_str) + 0.8)
                ws_sdrrffe_row += 1
            ws_sdrrffe.set_column(tech_col + ws_sdrrffe_col, tech_col + ws_sdrrffe_col, col_width)
        ws_sdrrffe_col += 8
        ws_sdrrffe.set_column(ws_sdrrffe_col, ws_sdrrffe_col, 3)
        ws_sdrrffe_col += 1


#================ @ Concurrency Restriction ====================
child_concurrency = root.find("concurrency_restriction_exception_list")

if isinstance(child_concurrency, ET.Element):

    et_allow_list = child_concurrency.find('allowed_list')

    allow_list_sigpath_a = []
    allow_list_sigpath_b = []
    ws_Concurrency = wb.add_worksheet('Concurrency Restrictions')
    ws_Concurrency.freeze_panes(2, 0)
    ws_Concurrency.set_row(0, 25)
    row_allow_list = 0
    col_allow_list = 0

    if isinstance(et_allow_list, ET.Element):
        for group_et in et_allow_list:
            sigpath_a_et = group_et.find('sig_path_a')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text+' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            allow_list_sigpath_a.append(sigpath_a_str)

            sigpath_a_et = group_et.find('sig_path_b')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text + ' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            allow_list_sigpath_b.append(sigpath_a_str)

        col_width_allow_list_A = []
        col_width_allow_list_B = []
        ws_Concurrency.write(row_allow_list, col_allow_list, 'ALLOWED LIST', format1)
        row_allow_list += 1
        col_width_allow_list_A.append(calc_col_width_by_str('ALLOWED LIST') + 0.8)
        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path A', format2)
        row_allow_list += 1
        col_width_allow_list_A.append(calc_col_width_by_str('Sig Path A') + 0.8)

        for list_num in range(0, len(allow_list_sigpath_a)):
            ws_Concurrency.write(row_allow_list, col_allow_list, allow_list_sigpath_a[list_num], format3)
            row_allow_list += 1
            col_width_allow_list_A.append(calc_col_width_by_str(allow_list_sigpath_a[list_num]) + 0.8)
        ws_Concurrency.set_column(col_allow_list, col_allow_list,adjust_concurrency_number_str_col_width(max(col_width_allow_list_A)))
        col_allow_list += 1

        row_allow_list = 1
        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path B', format2)
        row_allow_list += 1
        col_width_allow_list_B.append(calc_col_width_by_str('Sig Path B') + 0.8)

        for list_num in range(0, len(allow_list_sigpath_b)):
            ws_Concurrency.write(row_allow_list, col_allow_list, allow_list_sigpath_b[list_num], format3)
            row_allow_list += 1
            col_width_allow_list_B.append(calc_col_width_by_str(allow_list_sigpath_b[list_num]) + 0.8)

        ws_Concurrency.set_column(col_allow_list, col_allow_list,adjust_concurrency_number_str_col_width(max(col_width_allow_list_B)))

        col_allow_list += 1
        ws_Concurrency.set_column(col_allow_list, col_allow_list,3)
        col_allow_list += 1

    et_disallow_list = child_concurrency.find('disallowed_list')

    disallow_list_sigpath_a = []
    disallow_list_sigpath_b = []

    if isinstance(et_disallow_list, ET.Element):

        for group_et in et_disallow_list:
            sigpath_a_et = group_et.find('sig_path_a')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text+' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            disallow_list_sigpath_a.append(sigpath_a_str)

            sigpath_a_et = group_et.find('sig_path_b')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text + ' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            disallow_list_sigpath_b.append(sigpath_a_str)

        row_allow_list = 0
        col_width_disallow_list_A = []
        col_width_disallow_list_B = []
        ws_Concurrency.write(row_allow_list, col_allow_list, 'DISALLOWED LIST', format1)
        row_allow_list += 1
        col_width_disallow_list_A.append(calc_col_width_by_str('DISALLOWED LIST') + 0.8)

        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path A', format2)
        row_allow_list += 1
        col_width_disallow_list_A.append(calc_col_width_by_str('Sig Path A') + 0.8)

        for list_num in range(0, len(disallow_list_sigpath_a)):
            ws_Concurrency.write(row_allow_list, col_allow_list, disallow_list_sigpath_a[list_num], format3)
            row_allow_list += 1
            col_width_disallow_list_A.append(calc_col_width_by_str(disallow_list_sigpath_a[list_num]) + 0.8)
        ws_Concurrency.set_column(col_allow_list, col_allow_list,adjust_concurrency_number_str_col_width(max(col_width_disallow_list_A)))
        col_allow_list += 1

        row_allow_list = 1
        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path B', format2)
        row_allow_list += 1
        col_width_disallow_list_B.append(calc_col_width_by_str('Sig Path B') + 0.8)

        for list_num in range(0, len(disallow_list_sigpath_b)):
            ws_Concurrency.write(row_allow_list, col_allow_list, disallow_list_sigpath_b[list_num], format3)
            row_allow_list += 1
            col_width_disallow_list_B.append(calc_col_width_by_str(disallow_list_sigpath_b[list_num]) + 0.8)

        ws_Concurrency.set_column(col_allow_list, col_allow_list, adjust_concurrency_number_str_col_width(max(col_width_disallow_list_B)))

        col_allow_list += 1
        ws_Concurrency.set_column(col_allow_list, col_allow_list, 3)
        col_allow_list += 1

    et_msim_allow_list = child_concurrency.find('msim_allowed_list')

    mism_allow_list_sigpath_a = []
    mism_allow_list_sigpath_b = []

    if isinstance(et_msim_allow_list, ET.Element):
        for group_et in et_msim_allow_list:
            sigpath_a_et = group_et.find('sig_path_a')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text+' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            mism_allow_list_sigpath_a.append(sigpath_a_str)

            sigpath_a_et = group_et.find('sig_path_b')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text + ' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            mism_allow_list_sigpath_b.append(sigpath_a_str)

        row_allow_list = 0
        col_width_msim_allow_list_A = []
        col_width_msim_allow_list_B = []
        ws_Concurrency.write(row_allow_list, col_allow_list, 'MSIM ALLOWED LIST', format1)
        row_allow_list += 1
        col_width_msim_allow_list_A.append(calc_col_width_by_str('MSIM ALLOWED LIST') + 0.8)

        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path A', format2)
        row_allow_list += 1
        col_width_msim_allow_list_A.append(calc_col_width_by_str('Sig Path A') + 0.8)

        for list_num in range(0, len(mism_allow_list_sigpath_a)):
            ws_Concurrency.write(row_allow_list, col_allow_list, mism_allow_list_sigpath_a[list_num], format3)
            row_allow_list += 1
            col_width_msim_allow_list_A.append(calc_col_width_by_str(mism_allow_list_sigpath_a[list_num]) + 0.8)
        ws_Concurrency.set_column(col_allow_list, col_allow_list,adjust_concurrency_number_str_col_width(max(col_width_msim_allow_list_A)))
        col_allow_list += 1

        row_allow_list = 1
        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path B', format2)
        row_allow_list += 1
        col_width_msim_allow_list_B.append(calc_col_width_by_str('Sig Path B') + 0.8)

        for list_num in range(0, len(mism_allow_list_sigpath_b)):
            ws_Concurrency.write(row_allow_list, col_allow_list, mism_allow_list_sigpath_b[list_num], format3)
            row_allow_list += 1
            col_width_msim_allow_list_B.append(calc_col_width_by_str(mism_allow_list_sigpath_b[list_num]) + 0.8)

        ws_Concurrency.set_column(col_allow_list, col_allow_list,
                                  adjust_concurrency_number_str_col_width(max(col_width_msim_allow_list_B)))

        col_allow_list += 1
        ws_Concurrency.set_column(col_allow_list, col_allow_list, 3)
        col_allow_list += 1

    et_msim_disallow_list = child_concurrency.find('msim_disallowed_list')

    mism_disallow_list_sigpath_a = []
    mism_disallow_list_sigpath_b = []

    if isinstance(et_msim_disallow_list, ET.Element):
        for group_et in et_msim_disallow_list:
            sigpath_a_et = group_et.find('sig_path_a')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text+' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            mism_disallow_list_sigpath_a.append(sigpath_a_str)

            sigpath_a_et = group_et.find('sig_path_b')
            sigpath_a_str = ''
            for sigpath_et in sigpath_a_et:
                sigpath_a_str += (sigpath_et.text + ' ')
            sigpath_a_str = sigpath_a_str[0:len(sigpath_a_str) - 1]
            mism_disallow_list_sigpath_b.append(sigpath_a_str)

        row_allow_list = 0
        col_width_msim_disallow_list_A = []
        col_width_msim_disallow_list_B = []
        ws_Concurrency.write(row_allow_list, col_allow_list, 'MSIM DISALLOWED LIST', format1)
        row_allow_list += 1
        col_width_msim_disallow_list_A.append(calc_col_width_by_str('MSIM DISALLOWED LIST') + 0.8)

        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path A', format2)
        row_allow_list += 1
        col_width_msim_disallow_list_A.append(calc_col_width_by_str('Sig Path A') + 0.8)

        for list_num in range(0, len(mism_disallow_list_sigpath_a)):
            ws_Concurrency.write(row_allow_list, col_allow_list, mism_disallow_list_sigpath_a[list_num], format3)
            row_allow_list += 1
            col_width_msim_disallow_list_A.append(calc_col_width_by_str(mism_disallow_list_sigpath_a[list_num]) + 0.8)
        ws_Concurrency.set_column(col_allow_list, col_allow_list,adjust_concurrency_number_str_col_width(max(col_width_msim_disallow_list_A)))
        col_allow_list += 1

        row_allow_list = 1
        ws_Concurrency.write(row_allow_list, col_allow_list, 'Sig Path B', format2)
        row_allow_list += 1
        col_width_msim_disallow_list_B.append(calc_col_width_by_str('Sig Path B') + 0.8)

        for list_num in range(0, len(mism_disallow_list_sigpath_b)):
            ws_Concurrency.write(row_allow_list, col_allow_list, mism_disallow_list_sigpath_b[list_num], format3)
            row_allow_list += 1
            col_width_msim_disallow_list_B.append(calc_col_width_by_str(mism_disallow_list_sigpath_b[list_num]) + 0.8)

        ws_Concurrency.set_column(col_allow_list, col_allow_list,
                                  adjust_concurrency_number_str_col_width(max(col_width_msim_disallow_list_B)))

        col_allow_list += 1
        ws_Concurrency.set_column(col_allow_list, col_allow_list, 3)
        col_allow_list += 1

#================ @ Signal Path Selection Table ====================
child_signalpathselection = root.find("signal_path_selection_list_v2")
sigpath_sel_row = 0
sigpath_sel_col = 0
max_row_number = 0
if isinstance(child_signalpathselection, ET.Element):
    ws_sigpath_sel = wb.add_worksheet('Signal Path Selection Table')
    ws_sigpath_sel.freeze_panes(2, 0)
    ws_sigpath_sel.set_row(0, 25)
    sigpath_group_line = ['TX Group 0', 'TX Group 1', 'RX Group 0', 'RX Group 1', 'RX Group 2', 'RX Group 3',
                          'RX Group 4', 'RX Group 5', 'RX Group 6', 'RX Group 7']

    tech_sel_et = child_signalpathselection.find('sig_path_sel_lte_group')
    if isinstance(tech_sel_et, ET.Element):
        trx_group = get_sigpath_sel_trx_group_str(tech_sel_et)
        ws_sigpath_sel.merge_range(0, sigpath_sel_col, 0, sigpath_sel_col+9,'LTE Signal Path Selection Group', format1)
        for tech_col in range(0,10):
            sigpath_sel_row = 1
            col_width = calc_col_width_by_str(sigpath_group_line[tech_col])
            ws_sigpath_sel.write(sigpath_sel_row, tech_col+sigpath_sel_col, sigpath_group_line[tech_col], format2)
            sigpath_sel_row += 1
            for group_str in trx_group[tech_col]:
                ws_sigpath_sel.write(sigpath_sel_row, tech_col + sigpath_sel_col, group_str, format3)
                if calc_col_width_by_str(group_str) > col_width:
                    col_width = calc_col_width_by_str(group_str)
                sigpath_sel_row += 1
            ws_sigpath_sel.set_column(tech_col + sigpath_sel_col, tech_col + sigpath_sel_col, col_width)

        # ws_sigpath_sel.autofilter(1, sigpath_sel_col, sigpath_sel_row-1, sigpath_sel_col + 9)
        max_row_number = sigpath_sel_row
        sigpath_sel_col += 10
        ws_sigpath_sel.set_column(sigpath_sel_col, sigpath_sel_col, 3)
        sigpath_sel_col += 1

    tech_sel_et = child_signalpathselection.find('sig_path_sel_nr5g_group')
    if isinstance(tech_sel_et, ET.Element):
        trx_group = get_sigpath_sel_trx_group_str(tech_sel_et)
        ws_sigpath_sel.merge_range(0, sigpath_sel_col, 0, sigpath_sel_col + 9, 'NR5G Signal Path Selection Group',
                                   format1)
        for tech_col in range(0,10):
            sigpath_sel_row = 1
            col_width = calc_col_width_by_str(sigpath_group_line[tech_col])
            ws_sigpath_sel.write(sigpath_sel_row, tech_col+sigpath_sel_col, sigpath_group_line[tech_col], format2)
            sigpath_sel_row += 1
            for group_str in trx_group[tech_col]:
                ws_sigpath_sel.write(sigpath_sel_row, tech_col + sigpath_sel_col, group_str, format3)
                if calc_col_width_by_str(group_str) > col_width:
                    col_width = calc_col_width_by_str(group_str)
                sigpath_sel_row += 1
            ws_sigpath_sel.set_column(tech_col + sigpath_sel_col, tech_col + sigpath_sel_col, col_width)
        # ws_sigpath_sel.autofilter(1, sigpath_sel_col, sigpath_sel_row - 1, sigpath_sel_col + 9)
        if sigpath_sel_row > max_row_number:
            max_row_number = sigpath_sel_row
        sigpath_sel_col += 10
        ws_sigpath_sel.set_column(sigpath_sel_col, sigpath_sel_col, 3)
        sigpath_sel_col += 1

    tech_sel_et = child_signalpathselection.find('sig_path_sel_lte_nr5g_group')
    if isinstance(tech_sel_et, ET.Element):
        trx_group = get_sigpath_sel_trx_group_str(tech_sel_et)
        ws_sigpath_sel.merge_range(0, sigpath_sel_col, 0, sigpath_sel_col + 9, 'LTE/NR5G Signal Path Selection Group',
                                   format1)
        for tech_col in range(0, 10):
            sigpath_sel_row = 1
            col_width = calc_col_width_by_str(sigpath_group_line[tech_col])
            ws_sigpath_sel.write(sigpath_sel_row, tech_col + sigpath_sel_col, sigpath_group_line[tech_col], format2)
            sigpath_sel_row += 1
            for group_str in trx_group[tech_col]:
                ws_sigpath_sel.write(sigpath_sel_row, tech_col + sigpath_sel_col, group_str, format3)
                if calc_col_width_by_str(group_str) > col_width:
                    col_width = calc_col_width_by_str(group_str)
                sigpath_sel_row += 1
            ws_sigpath_sel.set_column(tech_col + sigpath_sel_col, tech_col + sigpath_sel_col, col_width)
        # ws_sigpath_sel.autofilter(1, sigpath_sel_col, sigpath_sel_row - 1, sigpath_sel_col + 9)
        if sigpath_sel_row > max_row_number:
            max_row_number = sigpath_sel_row
        sigpath_sel_col += 10
        ws_sigpath_sel.set_column(sigpath_sel_col, sigpath_sel_col, 3)
        sigpath_sel_col += 1
    if sigpath_sel_col > 1:
        ws_sigpath_sel.autofilter(1, 0, max_row_number - 1, sigpath_sel_col -1 )

#================ @ Ant Path Restriction Table ====================
child_antpath_restriction = root.find("antenna_restriction_exception_list")

if isinstance(child_antpath_restriction, ET.Element):
    ws_antpath_restriction = wb.add_worksheet('Ant Switch Path Restrictions')
    ws_antpath_restriction.freeze_panes(2, 0)
    ws_antpath_restriction.set_row(0, 25)
    antpath_restriction_sel_row = 0
    antpath_restriction_sel_col = 0
    antpath_res_group_line = ['Antenna Switch Path A', 'Antenna Switch Path B']

    allow_disallow_et = child_antpath_restriction.find('allowed_list')
    if isinstance(allow_disallow_et, ET.Element):
        antpath_restriction = get_antpath_restriction_tbl(allow_disallow_et)
        ws_antpath_restriction.merge_range(0, antpath_restriction_sel_col, 0, antpath_restriction_sel_col + 1,
                                           'Antenna Switch Path Allowed List',
                                           format1)
        for tech_col in range(0, 2):
            antpath_restriction_sel_row = 1
            col_width = calc_col_width_by_str(antpath_res_group_line[tech_col])
            ws_antpath_restriction.write(antpath_restriction_sel_row, tech_col + antpath_restriction_sel_col,
                                         antpath_res_group_line[tech_col], format2)

            antpath_restriction_sel_row += 1
            for group_str in antpath_restriction[tech_col]:
                ws_antpath_restriction.write(antpath_restriction_sel_row, tech_col + antpath_restriction_sel_col,
                                             group_str, format3)
                if calc_col_width_by_str(group_str) > col_width:
                    col_width = calc_col_width_by_str(group_str)
                antpath_restriction_sel_row += 1
            ws_antpath_restriction.set_column(tech_col + antpath_restriction_sel_col,
                                              tech_col + antpath_restriction_sel_col, col_width)
        antpath_restriction_sel_col += 2
        ws_antpath_restriction.set_column(antpath_restriction_sel_col, antpath_restriction_sel_col, 3)
        antpath_restriction_sel_col += 1

    allow_disallow_et = child_antpath_restriction.find('disallowed_list')
    if isinstance(allow_disallow_et, ET.Element):
        antpath_restriction = get_antpath_restriction_tbl(allow_disallow_et)
        ws_antpath_restriction.merge_range(0, antpath_restriction_sel_col, 0, antpath_restriction_sel_col + 1, 'Antenna Switch Path Dis-allowed List',
                                   format1)
        for tech_col in range(0,2):
            antpath_restriction_sel_row = 1
            col_width = calc_col_width_by_str(antpath_res_group_line[tech_col])
            ws_antpath_restriction.write(antpath_restriction_sel_row, tech_col+antpath_restriction_sel_col, antpath_res_group_line[tech_col], format2)

            antpath_restriction_sel_row += 1
            for group_str in antpath_restriction[tech_col]:
                ws_antpath_restriction.write(antpath_restriction_sel_row, tech_col + antpath_restriction_sel_col, group_str, format3)
                if calc_col_width_by_str(group_str) > col_width:
                    col_width = calc_col_width_by_str(group_str)
                antpath_restriction_sel_row += 1
            ws_antpath_restriction.set_column(tech_col + antpath_restriction_sel_col, tech_col + antpath_restriction_sel_col, col_width)
        antpath_restriction_sel_col += 2
        ws_antpath_restriction.set_column(antpath_restriction_sel_col, antpath_restriction_sel_col, 3)
        antpath_restriction_sel_col += 1

#================ @ Get Ant Path selsction Table ====================
child_band_class = root.find("band_classification_list")

antpath_select_sel_row = 0
antpath_select_sel_col = 0
is_antpath_sel_page_created = 0
if isinstance(child_band_class, ET.Element):
    ws_antpath_selection = wb.add_worksheet('Ant Switch Path Selection')
    ws_antpath_selection.freeze_panes(2, 0)
    ws_antpath_selection.set_row(0, 25)
    is_antpath_sel_page_created = 1

    band_class_group = [[], []]
    band_class_list = child_band_class.findall('band_class')
    for band_class_et in band_class_list:
        band_class_group[0].append(band_class_et.attrib['bandclass_name'])
        band_config_list = band_class_et.findall('band_config')
        band_config_s = ''
        for band_config_et in band_config_list:
            band_config_s += band_config_et.find('band').text + ' '
        band_config_s = band_config_s[0:len(band_config_s)-1]
        band_class_group[1].append(band_config_s)

    ws_antpath_selection.merge_range(0, antpath_select_sel_col, 0, antpath_select_sel_col + 1,
                                       'Band Classification Table',
                                       format1)
    band_class_title_line = ['Band Group Name', 'Band List']

    for tech_col in range(0, 2):
        antpath_select_sel_row = 1
        col_width = calc_col_width_by_str(band_class_title_line[tech_col])
        ws_antpath_selection.write(antpath_select_sel_row, tech_col + antpath_select_sel_col,
                                     band_class_title_line[tech_col], format2)

        antpath_select_sel_row += 1
        for group_str in band_class_group[tech_col]:
            ws_antpath_selection.write(antpath_select_sel_row, tech_col + antpath_select_sel_col,
                                         group_str, format3)
            if calc_col_width_by_str(group_str) > col_width:
                col_width = calc_col_width_by_str(group_str)
            antpath_select_sel_row += 1
        ws_antpath_selection.set_column(tech_col + antpath_select_sel_col,
                                          tech_col + antpath_select_sel_col, col_width)
    antpath_select_sel_col += 2
    ws_antpath_selection.set_column(antpath_select_sel_col, antpath_select_sel_col, 3)
    antpath_select_sel_col += 1

child_antpath_sel = root.find("ant_path_selection_list")

if isinstance(child_antpath_sel, ET.Element):
    if is_antpath_sel_page_created == 0:
        ws_antpath_selection = wb.add_worksheet('Ant Switch Path Selection')
        ws_antpath_selection.freeze_panes(2, 0)
        ws_antpath_selection.set_row(0, 25)

    antpath_sel_group_list = child_antpath_sel.findall('group')
    antpath_sel_group_s = [[], [],
                           [], [], [], [], [], [], [], [],
                           []]
    for antpath_sel_group_et in antpath_sel_group_list:
        tx_group_list = antpath_sel_group_et.findall('tx_operation')
        group_index = 0
        for tx_group_et in tx_group_list:
            ant_group_str = get_antpath_sel_id_str(tx_group_et, 'tx')
            antpath_sel_group_s[group_index].append(ant_group_str)
            group_index += 1

        for tx_num in range(group_index, 2):
            antpath_sel_group_s[tx_num].append('')

        rx_group_list = antpath_sel_group_et.findall('rx_operation')
        group_index = 2
        for rx_group_et in rx_group_list:
            ant_group_str = get_antpath_sel_id_str(rx_group_et, 'rx')
            antpath_sel_group_s[group_index].append(ant_group_str)
            group_index += 1

        for rx_num in range(group_index, 10):
            antpath_sel_group_s[rx_num].append('')

        if isinstance(antpath_sel_group_et.find('comments'), ET.Element):
            antpath_sel_group_s[10].append(antpath_sel_group_et.find('comments').text) # comment
        else:
            antpath_sel_group_s[10].append('')  # comment

    antpath_select_sel_row_bandclass = antpath_select_sel_row
    ws_antpath_selection.merge_range(0, antpath_select_sel_col, 0, antpath_select_sel_col + 9,
                                     'Antenna Selection Group Table',
                                     format1)
    if antpath_select_sel_row > 3:
        ws_antpath_selection.merge_range(2, antpath_select_sel_col, antpath_select_sel_row - 1, antpath_select_sel_col + 10, ' ', format2)
    antpath_sel_title_line = ['TX(PCC)', 'TX(SCC 1)',
                             'RX(PCC)', 'RX(SCC 1)', 'RX(SCC 2)', 'RX(SCC 3)', 'RX(SCC 4)', 'RX(SCC 5)', 'RX(SCC 6)', 'RX(SCC 7)',
                             'Comments']
    for tech_col in range(0, 11):
        antpath_select_sel_row = antpath_select_sel_row_bandclass
        col_width = calc_col_width_by_str(antpath_sel_title_line[tech_col])
        ws_antpath_selection.write(1, tech_col + antpath_select_sel_col,
                                     antpath_sel_title_line[tech_col], format2) # 总是在第二行写入
        # antpath_select_sel_row += 1
        for group_str in antpath_sel_group_s[tech_col]:
            ws_antpath_selection.write(antpath_select_sel_row, tech_col + antpath_select_sel_col,
                                       group_str, format3)
            if calc_col_width_by_str(group_str) > col_width:
                col_width = calc_col_width_by_str(group_str)
            antpath_select_sel_row += 1
        ws_antpath_selection.set_column(tech_col + antpath_select_sel_col,
                                        tech_col + antpath_select_sel_col, col_width)

    antpath_select_sel_col += 11
    ws_antpath_selection.set_column(antpath_select_sel_col, antpath_select_sel_col, 3)
    antpath_select_sel_col += 1





wb.close() # save to xlsx file