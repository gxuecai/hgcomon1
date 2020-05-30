import xml.etree.ElementTree as ET
import types
import ca_combo_class

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
print(rfc_path) # C:\CODE\MPSS.HI.1.0.c8-00198\modem_proc\rf\rfc_himalaya\common\etc\rf_card\rfc_Global_SDRV300_BoardID2_ag.xml

# file element tree and get root node
tree= ET.ElementTree()
tree.parse(rfc_path)
root=tree.getroot()

# root RFC
print('root type', type(root))
print('root tag','=========',root.tag,'=========')
print('root length:', len(root))
assert isinstance(root, ET.Element)

# get ca_combo_list node
child=root.find("ca_combo_list")
print_tree_element_info(child,'child')
assert isinstance(child, ET.Element)

# get combo_group node--
child_child= child[0]
print_tree_element_info(child_child,'child_child')
assert isinstance(child_child, ET.Element)

# get ca_4g_combos node
child_child_child=child_child.find("ca_4g_combos")
print_tree_element_info(child_child_child,'child_child_child')
assert isinstance(child_child_child, ET.Element)

# handle 4g combos
print_tree_element_info(child_child_child[0],'child_child_child[0]')
lte_combo_list=[]
for ca_4g_combos in child_child_child:
    #assert isinstance(ca_4g_combos, ET.Element)
    print(ca_4g_combos.text)
    lte_combo_list.append(ca_combo_class.Lte_ca_combo(ca_4g_combos.text))

print(len(lte_combo_list))



