import os
from xml_parser_sysband import system_band_dic

os.system("cls") # cmd window clear screen

max_sysband_enum=tuple(system_band_dic.keys())[-1]
min_sysband_enum=tuple(system_band_dic.keys())[1]

print("\n\n-----------------START-----------------\n")

while 1:
    try:
        sysband_num = input("Input system band enum(%s ~ %s): " % (min_sysband_enum, max_sysband_enum))
        if sysband_num == 'exit':
            break
        print('Sys Band Enum: %s ===> Tech Band: %s \n' %(sysband_num,system_band_dic[sysband_num]))
    except:
        print("Sys Band Enum Input Invalid, Please double check the enum in the log prints \n")


end_string='''
------------------END------------------
'''
print(end_string)
os.system("pause") # dos cmd window pause