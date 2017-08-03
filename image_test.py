import xml.etree.ElementTree as etree

tree = etree.parse('slide1.xml')
root = tree.getroot()
print( root.tag)
print root.attrib
#hardcoding the namespace
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
#print p
"""
for child in root:
	print child.tag
"""
spTree =tree.find('.//'+p+'spTree')
#print node.tag

#for child in node:
#	print child
#node = node.find(p+"spTree")

for child in spTree:
	print child.tag

print
rect_list = spTree.findall(p+'sp')

for child in rect_list:
	sp = child.find('.//'+p+'cNvPr')
	print 'id:' +sp.get('id') + " | name:"+sp.get('name')
