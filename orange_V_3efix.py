#mapdatain_5.py
#to read in ALL data
#TRYING TO FIGURE OUT WHERE PHONY FIGURES COMING FROM

"""
These 3 imports needed for the full program
# from classes import *
# from code import *
"""
import math
import matplotlib.pyplot as plt
from scipy import stats
import numpy as np
import random
import time
# import data
import data as pass_data
from orangeclass_V_3efix import * #import the FUNCTIONS from the CLASSES
# from boxesmap3_1 import *


def filein (fname):
    numlines=0
    xin=[]
    f=open(fname,'r')
    for line in f:
        #print (line, end='')
        xin.append(line)
        numlines=numlines+1
    f.close()
    return xin,numlines

def fileout (filename,filedata):
    f2=open(filename,'w')
    f2.write(filedata)
    f2.close()

def getxy (fname):
    data,numlines=filein(fname)
#    dataline=['0' for i in range(numlines)]
    x=['0' for i in range(numlines)]
    y=['0' for i in range(numlines)]
    for i in range(numlines):
        xline=data[i]
        xline2=xline.split('\t')
        xline2[-1]=xline2[-1].replace('\n','')
        #print ('\nxline2',xline2)
        x[i]=eval(xline2[0])
        y[i]=eval(xline2[1])
    return x,y,numlines

def getx (fname):
    data,numlines=filein(fname)
    print ('\ndata\n',data,'\nlines',numlines)
    x=['0' for i in range(numlines)]
    for i in range(numlines):
        x[i]=data[i].replace('\n','')
        x[i]=eval(x[i])
    return x,numlines

def getxn(fname):
    data,numlines=filein(fname)
#print (numlines)
#print (data)
    dataline=['0' for i in range(numlines)]
    for i in range(numlines):
        x=data[i]
        y=x.split('\t')
        y[-1]=y[-1].replace('\n','')
        dataline[i]=y
    #print ('\n\nascii-input',dataline)
    xdata=dataline[:]
    print (xdata)
    for i in range (numlines):
        inline=len(dataline[i])
        for j in range (inline):
            if xdata[i][j] != '':
                xdata[i][j]=eval(xdata[i][j])
            else:
                xdata[i][j]=None
    #print ('dataline',dataline)
    #print ('xdata',xdata)
    return xdata, numlines
    
    
#MIDIFYING THIS FUNCTION TO GET THE COLOR
def getxnsecondstring(fname):   #get n inputs from each line
#first column=text, next colu1mns all numbers
#here: variable name, x, y, height, width
    data,numlines=filein(fname)
#print (numlines)
#print (data)
    dataline=['0' for i in range(numlines)]
    for i in range(numlines):
        x=data[i]
        y=x.split('\t')
        y[-1]=y[-1].replace('\n','')
        dataline[i]=y
    #print ('\n\nascii-input',dataline)
    xdata=dataline[:]
    for i in range (numlines):
        inline=len(dataline[i])
        for j in range (2,inline):
            if xdata[i][j] != '':
                xdata[i][j]=eval(xdata[i][j])
            else:
                xdata[i][j]=None8
                
    #print ('dataline',dataline)
    #print ('xdata',xdata)
    return xdata, numlines

#-------------------------------------------------------------------------
def lslin(invars,invar):
    #print('\ncurrent value of ',invars,' is= ',invar)
    outvars=input('\nchange to (def=no change)')
    if (outvars==''):
        return invar
    else:
        outvar=eval(outvars)
    return outvar
    
#END OF DEFINED FUNCTIONS---------------


#START DATA INPUT
#give it just a number n, will find files cn.txt, bn.txt, mn.txt, icn.txt, btext and bxy
fast=input('\n ONLY NUMBER n and I will find cn.txt, etc. (#/a, Def=a)')
if fast.isdigit():
    fnamec='c'+fast+'.txt'
    fnameb='b'+fast+'.txt'
    fnamem='m'+fast+'.txt'
    fnameic='ic'+fast+'.txt'
    fnamebtextbxy='btextbxy'+fast+'.txt'
else:
    fname=input('\nfilename for array c [I will add .txt]=  ')
    fnamec=fname+'.txt'
    
    fname=input('\nfilename for array b [I will add .txt]=  ')
    fnameb=fname+'.txt'
    
    fname=input('\nfilename for array m [I will add .txt]=  ')
    fnamem=fname+'.txt'
    
    fname=input('\nfilename for array IC [I will add .txt]=  ')
    fnameic=fname+'.txt'
    
    fname=input('\nfilename for bxy, btext [I will add .txt]=  ')
    fnamebtextbxy=fname+'.txt'

#get the files
c,numc=getxn(fnamec)
b,numb=getx(fnameb)
m,numm=getx(fnamem)
ic,numic=getx(fnameic)
btextbxydata,numvar=getxnsecondstring(fnamebtextbxy)

#check for consistentcy
if (numc**4!=numb*numm*numic*numvar):
    print ("\nFATAL WARNING - input issue - numbers c,b,m,ic,bxy,btext don't match")
    quit()

#PART ONE make original m, b, c, ic arrays (NOT matrices) and print
ma=np.array(m)
ba=np.array(b)
ca=np.array(c)
ica=np.array(ic)
print ('\nca= ',ca)
print ('\nba= ',ba)
print ('\nma= ',ma)
print ('\nica= ',ica)

change=input('\nWant to CHANGE parameters (y/n), def=n')
if (change=='y' or change=='Y'):
    c=lslin('c',c)
    b=lslin('b',b)
    m=lslin('m',m)
    ic=lslin('ic',ic)
    
    ma=np.array(m)
    ba=np.array(b)
    ca=np.array(c)
    ica=np.array(ic)
    print ('\n\nNEW PARAMTER VALUES ARE:')
    print ('\nca= ',ca)
    print ('\nba= ',ba)
    print ('\nma= ',ma)
    print ('\nic= ',ica)
else:
    pass

#PART TWO read in the variable names and box locations in the plot btext and bxy(x,y,h,w)
print('\n numvar(from btextbxy)= ',numvar)
print('\n btextbxydata= ',btextbxydata)

#COMPUTE (x,y)=[0,1] needed from PPTX
bx=[btextbxydata[i][2] for i in range (numvar)]
by=[btextbxydata[i][3] for i in range (numvar)]
wx=[btextbxydata[i][5] for i in range (numvar)]
hy=[btextbxydata[i][4] for i in range (numvar)]

#note this scaling has changed 2017-07-06
#SCALE as needed for the plot
xp=[0. for i in range(numvar)]
yp=[0. for i in range(numvar)]
xp2=[0. for i in range(numvar)]
yp2=[0. for i in range(numvar)]
for i in range(numvar):
    xp[i]=(bx[i] + 0.5*wx[i])
    yp[i]=(by[i] + 0.5*hy[i])
maxx,minx=max(xp),min(xp)
maxy,miny=max(yp),min(yp)
for i in range(numvar):
    xp2[i]=0.9*(xp[i]-minx)/(maxx-minx)+0.05
    yp2[i]=1-(0.9*(yp[i]-miny)/(maxy-miny)+0.05)
    
bxy=[[xp2[i],yp2[i]] for i in range(numvar)]
print ('\nbxy=  ',bxy)

#PARAEMTERS NEEDED FOR THE NUMERICAL INTEGRATION
dt=.001
numdata=30000
t=[0. for i in range(numdata)]
z=np.array([ica for i in range (numdata)])

#READY TO PASS ON DATA----------------------------------------------------------
#wrap parameters to pass into function 
pass_data.numdata=numdata
pass_data.ca=ca
pass_data.dt=dt
pass_data.ma=ma
pass_data.ba=ba
pass_data.numc=numc
pass_data.z=z
pass_data.t=t
pass_data.ica=ica
#NEW DATA's ADDED BELOW
pass_data.fnamec=fnamec
pass_data.fnamem=fnamem
pass_data.fnameb=fnameb
pass_data.fnamebtextbxy=fnamebtextbxy
pass_data.dt=dt
#EVEN MORE DATA'S ADDED BELOW
pass_data.a=ca
pass_data.bxy=bxy
pass_data.btext=str([btextbxydata[i][0] for i in range(numvar)])
pass_data.b=ica
pass_data.labels=[btextbxydata[i][0] for i in range(numvar)]
#ADDING BOX COLORS
pass_data.boxcolor=[btextbxydata[i][1] for i in range(numvar)]

#this stuff to call data_3.py and make PLOTS!

#FIRST LET'S CHECK INPUT THEN LATER DO THE CALL------------
zzz=App()

#FIRST LET'S CHECK INPUT THEN LATER DO THE CALL------------


# zzz.MakeWindow()
# zzz.MakeSample()



"""
THESE NEEDED FOR THE FULL PROGRAM
# App.recalculate(pass_data)
#---------->call App
# callGUI()
"""

