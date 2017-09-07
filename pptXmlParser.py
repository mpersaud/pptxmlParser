import xml.etree.ElementTree as etree
import numpy as np
import sys
#hardcoding the namespace
p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def getInput():
	input_slide = raw_input("Enter XML filename: ");
	if(input_slide.find('.xml')!=-1):
		input_slide= input_slide[:input_slide.find('.xml')]
	return input_slide
	
def getOutputFile(i):
	if(i==1):
		input_slide = raw_input("Enter nodes output filename: ");
	if (i==0):
		input_slide = raw_input("Enter matrix output filename: ");
	if(input_slide.find('.txt')!=-1):
		input_slide= input_slide[:input_slide.find('.txt')]
	return input_slide

###SCALE FACTOR = 12700
SCALE_FACTOR = 12700
MAX_WEIGHT = 0

input_slide=getInput()
tree = etree.parse(input_slide+".xml")
root = tree.getroot()

spTree =tree.find('.//'+p+'spTree')

shape_list = spTree.findall(p+'sp')
cxn_list =  spTree.findall(p+'cxnSp')
#debug purpose
i = 0
negative=0
positive =0
#counting nodes/rect
r=0;

##MATRIX && RECT/NODE MAP
matrix = [[]]
mapping = {}
print 'Running...'
nodes_file = open(getOutputFile(1)+".txt","w")
#SHAPES LIST
for child in shape_list:
	#non-visual properties
	shape = child.find('.//'+p+'cNvPr')
	#shape properties
	spPr = child.find(''+p+"spPr")

	#PRINTS OUT RECTANGLES WITH ID AND NAME, OFFSET , WIDTH , HEIGHT
	if child[1][1].attrib.get('prst') == "rect" or etree.iselement(spPr.find(a+"custGeom")) :


		xfrm = spPr.find(a+"xfrm")

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
		#debugging purpose
		#print "|"+'id:' +shape.get('id') + " | name:"+shape.get('name')+"| Rectangle:"+full_text + "| x_offset: " + x_offset + "| y_offset:" + y_offset + "| width:" + width + "| height:" + height
		nodes_file.write(full_text + " " + x_offset + " " + y_offset + " " + width + " " + height)
		nodes_file.write('\n')
		i=i+1
		#add to map and increment node counter
		mapping[int(shape.get('id'))]=r
		#print "Node "+str(r)+ ":"+full_text
		r=r+1



	# NOT NEEDED ANYMORE
	#ARROWS CURVED DOWN or CURVED LEFT or Plus and Minu with ID AND NAME , OFFSET , WIDTH , HEIGHT
	if child[1][1].attrib.get('prst') == "curvedDownArrow" or child[1][1].attrib.get('prst') == "curvedLeftArrow" or child[1][1].attrib.get('prst') == "mathPlus" or child[1][1].attrib.get('prst') == "mathMinus":
		#print child[1][1].attrib
		spPr = child.find(''+p+"spPr")
		xfrm = spPr.find(a+"xfrm")

		x_offset= xfrm.find(a+'off').attrib.get('x')
		y_offset= xfrm.find(a+'off').attrib.get('y')
		width= xfrm.find(a+'ext').attrib.get('cx')
		height= xfrm.find(a+'ext').attrib.get('cy')
		#print 'id:' +shape.get('id') + " | name:"+shape.get('name')+ "| x_offset: " + x_offset + "| y_offset:" + y_offset + "| width:" + width + "| height:" + height
		i=i+1

#close nodes_file
nodes_file.close()
#initalize the matrix
matrix = np.matrix([[0]*r]*r)

#CONNECTORS LIST consisting of connector shapes <cxnSp>
for child in cxn_list:
	CxnSpPr = child.find('.//'+p+'nvCxnSpPr')
	cNvPr = CxnSpPr.find('.//'+p+'cNvPr')
	spPr = child.find(''+p+"spPr")
	xfrm = spPr.find(a+"xfrm")
	aln = spPr.find(a+"ln")

	#DEFAULT as Positive
	RGB="+"
	asolidFill = aln.find(a+"solidFill")
	if(asolidFill!=None):
		aRGB = asolidFill.find(a+"srgbClr")
		if(aRGB!=None):
			RGB = aRGB.get('val')
			if(RGB=="FF0000"):
				RGB="-"
				negative=negative+1
		else:
			positive=positive+1
	else:
		RGB="+"
		positive=positive+1

	#DEAFULT IS SCALE_FACTOR if not grab value
	line_width = aln.get('w')
	if(line_width==None):
		line_width='12700'

	if float(line_width)>MAX_WEIGHT:
		MAX_WEIGHT=float(line_width)
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
	#print 'id:' +cNvPr.get('id') + "| Sign:"+RGB+"| name:"+cNvPr.get('name')+ "| x_offset: " + x_offset + "| y_offset:" + y_offset + "| width:" + width + "| height:" + height + "| Start Con:" + start_Cxn + "| End Con:" + end_Cxn +" | line_width: "+line_width
	#change to mapping and input to matrix
	i=i+1
	if start_Cxn =='0' or end_Cxn =='0':
		continue
	start= mapping.get(int(start_Cxn))
	end= mapping.get(int(end_Cxn))

	if(RGB=='-'):
		line_width=float(line_width)*-1.0

	matrix[end].put(start,(float(line_width)*1.0))


matrix =np.multiply(1.0/SCALE_FACTOR,matrix)
matrix =np.matrix(np.round(matrix,3))

output = open(getOutputFile(0)+".txt","w")
list = matrix.tolist()
for i in range(r):
	output.write((str(list[i])[1:-1]).replace(',',' '))
	output.write('\n')
output.close()

#output.write(str(matrix.A))
print 'Finished.'
def debug():
	print
	print "Objects: "+str(i)
	print "Nodes: "+str(r)
	print "Negative Connects: "+str(negative)
	print "Positive Connects: "+str(positive)
	print "Node Mapping to ID: "+str(sorted( ((v,k) for k,v in mapping.iteritems()), reverse=False))
	print "--------------------------"
	print "Directed In-Graph"
	print matrix
	print "MAX_WEIGHT:"+str(MAX_WEIGHT)
	print "-------------------------"
	print "Directed Out-Graph(Transpose)"

	print matrix.getT()
	print
#debug()
