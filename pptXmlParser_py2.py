import xml.etree.ElementTree as etree
import numpy as np

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
accent1 ="4F81BD"
accent2 ="C0504D"
accent3 ="9BBB59"
accent4 ="8064A2"
accent5 ="4BACC6"
accent6 ="F79646"

input_slide=getInput()
tree = etree.parse(input_slide+".xml")
root = tree.getroot()

spTree =tree.find('.//'+p+'spTree')

shape_list = spTree.findall(p+'sp')
cxn_list =  spTree.findall(p+'cxnSp')
#debug purpose
o = 0
negative=0
positive =0
#counting nodes/rect
nodeNum=0;

##MATRIX && RECT/NODE MAP
matrix = [[]]
mapping = {}
print ('Running...')
nodes_file = open(getOutputFile(1)+".txt","w")
#SHAPES LIST
for child in shape_list:
	#non-visual properties
	shape = child.find('.//'+p+'cNvPr')
	#shape properties
	spPr = child.find(''+p+"spPr")

	#PRINTS OUT RECTANGLES WITH ID AND NAME, OFFSET , WIDTH , HEIGHT
	if child[1][1].attrib.get('prst') == "rect" or etree.iselement(spPr.find(a+"custGeom")) :

		rectSolidFill= spPr.find(a+"solidFill")
		rectColor = rectSolidFill.find(a+"schemeClr")
		if(rectColor == None):
			rectColor = rectSolidFill.find(a+"srgbClr").get('val')
		else:
			rectColor = rectSolidFill.find(a+"schemeClr").get('val')

		xfrm = spPr.find(a+"xfrm")

		x_offset= xfrm.find(a+'off').attrib.get('x')
		y_offset= xfrm.find(a+'off').attrib.get('y')
		width= xfrm.find(a+'ext').attrib.get('cx')
		height= xfrm.find(a+'ext').attrib.get('cy')

		full_text = ""
		textBody = child.find(''+p+'txBody')
		t = textBody.findall('.//'+a+'t')
		for elem in t:
			full_text+="".join(elem.text)

		identifier = str(nodeNum+1)
		color = "yellow"
		if(rectColor == "accent2" or rectColor=="C0504D"):
			color = "gray"
		nodes_file.write(identifier+" " +full_text.rstrip() + "\t" + color + "\t"+x_offset + "\t" + y_offset + "\t" + width + "\t" + height)
		nodes_file.write('\n')
		o=o+1
		#add to map and increment node counter
		mapping[int(shape.get('id'))]=nodeNum
		#print "Node "+str(r)+ ":"+full_text
		nodeNum=nodeNum+1

#close nodes_file
nodes_file.close()
#initalize the matrix
matrix = np.matrix([[0]*nodeNum]*nodeNum)

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
		line_width= float('12700')*2

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
    #change to mapping and input to matrix
	o=o+1
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
for i in range(nodeNum):
	output.write((str(list[i])[1:-1]).replace(', ','\t'))
	output.write('\n')
output.close()

print ('Finished.')
def debug():
	print
	print ("Objects: "+str(o))
	print ("Nodes: "+str(nodeNum))
	print ("Negative Connects: "+str(negative))
	print ("Positive Connects: "+str(positive))
	print ("Node Mapping to ID: "+str(sorted( ((v,k) for k,v in mapping.items()), reverse=False)))
	print ("--------------------------")
	print ("Directed In-Graph")
	print (matrix)
	print ("MAX_WEIGHT:"+str(MAX_WEIGHT))
	print ("-------------------------")
	print ("Directed Out-Graph(Transpose)")

	print (matrix.getT())
	print
