import xml.etree.ElementTree as etree

tree = etree.parse('slide1.xml')
root = tree.getroot()
# print( root.tag)
# print root.attrib
#hardcoding the namespace
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
#print p

# for child in root:
# 	print child.tag, child.attrib, child.text
#

spTree =tree.find('.//'+p+'spTree')
#print node.tag

#for child in node:
#	print child
#node = node.find(p+"spTree")

#for child in spTree:
#	print child.tag

print
shape_list = spTree.findall(p+'sp')
cxn_list =  spTree.findall(p+'cxnSp')
#debug purpose
i = 0

#SHAPES LIST
for child in shape_list:
	#non-visual properties
	shape = child.find('.//'+p+'cNvPr')

	#print child.tag
	#print child[0][0].attrib
	#print child[1][1].attrib
	#print shape

	#PRINTS OUT RECTANGLES WITH ID AND NAME, OFFSET , WIDTH , HEIGHT
	if child[1][1].attrib.get('prst') == "rect":
	 	#print child[1][1].attrib
		spPr = child.find(''+p+"spPr")
		xfrm = spPr.find(a+"xfrm")
		#print xfrm.attrib
		x_offset= xfrm.find(a+'off').attrib.get('x')
		y_offset= xfrm.find(a+'off').attrib.get('y')
		width= xfrm.find(a+'ext').attrib.get('cx')
		height= xfrm.find(a+'ext').attrib.get('cy')

		full_text = ""
		textBody = child.find(''+p+'txBody')
		t = textBody.findall('.//'+a+'t')
		for elem in t:
			#print elem.text
			full_text+="".join(elem.text)
		print "|"+'id:' +shape.get('id') + " | name:"+shape.get('name')+"| Rectangle:"+full_text + "| x_offset: " + x_offset + "| y_offset:" + y_offset + "| width:" + width + "| height:" + height
		#print child[3][2][1][1].text
		i=i+1



	#ARROWS CURVED DOWN or CURVED LEFT or Plus and Minu with ID AND NAME , OFFSET , WIDTH , HEIGHT
	if child[1][1].attrib.get('prst') == "curvedDownArrow" or child[1][1].attrib.get('prst') == "curvedLeftArrow" or child[1][1].attrib.get('prst') == "mathPlus" or child[1][1].attrib.get('prst') == "mathMinus":
		#print child[1][1].attrib
		spPr = child.find(''+p+"spPr")
		xfrm = spPr.find(a+"xfrm")
		#print xfrm.attrib
		x_offset= xfrm.find(a+'off').attrib.get('x')
		y_offset= xfrm.find(a+'off').attrib.get('y')
		width= xfrm.find(a+'ext').attrib.get('cx')
		height= xfrm.find(a+'ext').attrib.get('cy')
		print 'id:' +shape.get('id') + " | name:"+shape.get('name')+ "| x_offset: " + x_offset + "| y_offset:" + y_offset + "| width:" + width + "| height:" + height
		i=i+1

#CONNECTORS LIST consisting of connector shapes <cxnSp>
for child in cxn_list:
	CxnSpPr = child.find('.//'+p+'nvCxnSpPr')
	cNvPr = CxnSpPr.find('.//'+p+'cNvPr')
	spPr = child.find(''+p+"spPr")
	xfrm = spPr.find(a+"xfrm")
	#print xfrm.attrib
	cNvCxnSpPr = CxnSpPr.find(p+'cNvCxnSpPr')
	start_Cxn = cNvCxnSpPr.find(a+'stCxn')
	if(etree.iselement(start_Cxn)):
		start_Cxn= start_Cxn.attrib.get('id')
	else:
		start_Cxn='0'


	end_Cxn = cNvCxnSpPr.find(a+'endCxn')
	if(etree.iselement(end_Cxn)):
		end_Cxn= end_Cxn.attrib.get('id')
	else:
		end_Cxn='0'

	x_offset= xfrm.find(a+'off').attrib.get('x')
	y_offset= xfrm.find(a+'off').attrib.get('y')
	width= xfrm.find(a+'ext').attrib.get('cx')
	height= xfrm.find(a+'ext').attrib.get('cy')
	print 'id:' +cNvPr.get('id') + " | name:"+cNvPr.get('name')+ "| x_offset: " + x_offset + "| y_offset:" + y_offset + "| width:" + width + "| height:" + height + "| Start Con:" + start_Cxn + "| End Con:" + end_Cxn
	i=i+1

print
print i
