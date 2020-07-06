import xml.etree.ElementTree as ET
import xlsxwriter as XW
import types
import re


def calc_col_width_by_str(write_str):
    str_list = re.split(r'\n', write_str)
    str_len_list = [len(item) for item in str_list]
    return max(str_len_list)


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
ws_signalpath = wb.add_worksheet('sigpaths')
ws_signalpath.freeze_panes(1, 0)
ws_signalpath.freeze_panes(1, 1)

format1 = wb.add_format({'font_size': 10, 'font_name': 'Calibri', 'bold': 1, 'font_color': 'white',
                         'fg_color': 'green',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1}) # 自动换行
format2 = wb.add_format({'font_size': 8, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': '#F2F2F2',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1, 'shrink': 0})
format3 = wb.add_format({'font_size': 8, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': 'white',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 1})
format4 = wb.add_format({'font_size': 8, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': '#F2F2F2',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 0, 'shrink': 1})
format5 = wb.add_format({'font_size': 8, 'font_name': 'Calibri', 'bold': 0, 'font_color': 'black',
                         'fg_color': 'white',
                         'bottom': 1, 'top': 1, 'right': 1, 'left': 1,
                         'align': 'center', 'valign': 'vcenter',
                         'text_wrap': 0, 'shrink': 1})

root=tree.getroot()

# root RFC
print('root type', type(root))
print('root tag','=========',root.tag,'=========')
print('root length:', len(root))
assert isinstance(root, ET.Element)

# get sig_paths node
child=root.find("sig_paths")
print_tree_element_info(child,'child')
assert isinstance(child, ET.Element)

Heading_list = ['Path ID', 'Type\n(rx/tx)', 'PRX?', 'Max\nTX BW', 'PWR\nClass', 'Functionality',
                'Cal Ref\nsigpath', 'Path\noverride\nIndex', 'MCS\n256QAM', 'Disabled', 'FBRX', 'Sigpath\nGroup',
                'Ant SW\npath', 'Split band\nsigpath Map', 'Split Band', 'Tech Bands', 'TRX', 'ELNA0', 'ELNA1', 'PA', 'PAPM', 'PAPM HUB',
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
            if ant_i < ant_num:
                ant_sw_path += ' '

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

wb.close() # save to xlsx file