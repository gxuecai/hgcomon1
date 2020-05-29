import xml.etree.ElementTree as ET
import types
'''
Python etree handle XML file
'''
tree= ET.ElementTree();
tree.parse('sys_band_enum.xml')
root=tree.getroot()

print('root type', type(root))
print('root tag','=========',root.tag,'=========')
print('root length:', len(root))
assert isinstance(root, ET.Element)

#children= root.findall('tech') # I think it's used for the scenario in which there are different 'tags' in same level tree

system_band_dic={0:'B1'}
child_i = 0
print('    child tag', '----------',root[0].tag,'----------')
for child in root:
    print('    child type', type(child))
    print('    child index %d length: %d' %(child_i, len(child)))
    assert isinstance(child, ET.Element)
    print('    child tag attrib',child.tag,child.attrib)
    child_i=child_i+1
    print('        child_child tag', '**********', child[0].tag, '***********')
    child_child_i=0
    for child_child in child:
        assert isinstance(child_child, ET.Element)
        print('        child_child type', type(child_child))
        print('        child_child index %d' % child_child_i)
        print('        child_child tag attrib', child_child.tag, child_child.attrib)
        child_child_i= child_child_i+1
        # save to system_band_dic
        temp_dic=child_child.attrib
        system_band_dic[temp_dic['enum_num']]=temp_dic['band']

print(system_band_dic)


