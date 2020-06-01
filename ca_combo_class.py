import re
# ---------------------------Header------------------------------
# define ca/endc combos class in this file
# ---------------------------------------------------------------


class Lte_ca_combo:

    def __init__(self, ca_string):
        # object instance variables which are also visible to other funcs of this class
        self.dl_ca_list = []
        self.ul_ca_list = []
        self.dl_band_num = 0
        self.ul_band_num = 0
        # parse the ca string from RFC to the object ca info variables
        self.parse_ca_list(ca_string)
        self.print_ca_info()

    # parse the ca band info, ant info
    def parse_ca_list(self, ca_string):
        # split the bands by '+'
        s_list = re.split(r'\+', ca_string)
        for ss_i in s_list:
            # match the DL band info and ant
            ss1_dl = re.search(r'B([0-9]+)[A-Z]\[([1-9])', ss_i)
            #print('band: %s, ant: %s' % (ss1_dl.group(1), ss1_dl.group(2)))
            self.dl_ca_list.append((ss1_dl.group(1), ss1_dl.group(2)))
            self.dl_band_num+=1
            # match the UL ant info
            ss1_ul = re.search(r';[A-F]\[([1-9])', ss_i)
            if ss1_ul:
                #print('ul ant', ss1_ul.group(1))
                self.ul_ca_list.append(ss1_dl.group(1))
                self.ul_band_num+=1

    # print self object instance variables of ca info
    def print_ca_info(self):
        print('dl_ca_list: ',self.dl_ca_list,' ul_ca_list:',self.ul_ca_list,' ca band number: ', self.dl_band_num)

aaa= Lte_ca_combo('B2A[2];A[1]+B46E[2,2,2,2]+B48A[4]')

'''
ss = re.split(r'\+', 'B2A[2];A[1]+B46E[2,2,2,2]+B48A[4]')
print(ss)

for ss_i in ss:
    ss1_dl = re.search(r'B([0-9]+)[A-Z]\[([1-9])', ss_i)
    print('band: %s, ant: %s' % (ss1_dl.group(1), ss1_dl.group(2)))
    ss1_ul = re.search(r';[A-F]\[([1-9])', ss_i)
    if ss1_ul:
        print('ul ant',ss1_ul.group(1))

'''


