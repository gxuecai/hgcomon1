import re

def get_subset_list(bandlist):
    ret_subset = []
    len_list = len(bandlist)
    if len_list == 1:
        return [bandlist]
    else:
        subset_fn_1 = get_subset_list(bandlist[1:len_list])
        ret_subset.append([bandlist[0]])
        ret_subset += subset_fn_1
        for subsets in subset_fn_1:
            ret_subset.append(subsets+[bandlist[0]])
        return ret_subset

def list_to_combo_string(combo_list):
    len_list = len(combo_list)
    combo_list.sort()
    ret_string = ''
    for i in range(0,len_list):
        if i == len_list - 1:
            ret_string += combo_list[i]
        else:
            ret_string += combo_list[i] + '+'
    return ret_string

def get_subset_combo_string(combo_string):

    subs = re.split(r'\+', combo_string)
    # print(subs)
    band_num = len(subs)
    sub_set_list = get_subset_list(subs)
    # print(sub_set_list)

    valid_sub_set_list = []
    for sub_set_list_i in sub_set_list:
        sub_set_list_i_string = list_to_combo_string(sub_set_list_i)
        if len(sub_set_list_i) == band_num:
            valid_sub_set_list.append(sub_set_list_i_string)
        else:
            ul_band_num = sub_set_list_i_string.count(';')
            is_ENDC_band = ('N' in sub_set_list_i_string) and ('B' in sub_set_list_i_string)

            if ((is_ENDC_band == True) and (ul_band_num >= 2)) or ((is_ENDC_band == False) and (ul_band_num >= 1)):
                valid_sub_set_list.append(sub_set_list_i_string)

    return valid_sub_set_list



'''
fo = open("log.txt", "w")

fo.write( "www.runoob.com!\nVery good site!\n")

fo.close()

subset_test = get_subset_list(['B1', 'B2','B3','N66'])
print(subset_test)
print(get_subset_combo_string('B1A[2];A[1]+B3A[4]+N78A[4];A[1]'))
print(get_subset_combo_string('B1A[2];A[1]+B3A[4];A[1]+N78A[4];A[1]'))
print(get_subset_combo_string('N1A[2];A[1]'))
print(get_subset_combo_string('B1A[2];A[1]'))
print(get_subset_combo_string('B1A[2];A[1]+B3A[4]+B78A[4];A[1]'))
# print(list_to_combo_string(['N78A[4];A[1]', 'B3A[4]', 'B1A[2];A[1]']))

'''