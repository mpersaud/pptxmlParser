"""
box_II_!.py
    create class of boxes
    readin bxy box locations
    plot boxes

NOW TRYINGING:
    1) RETURN CLICKS INSIDE BOXES
    2) INPUT NEW DATA
"""
#TRYING TO FIGURE OUT WHERE PHONY FIGURES COMING FROM


import tkinter as tk
import data
import numpy as np
from decimal import *
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import time
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg

class TextBox:
    list_box=[]
    selected_box_id=-1
    colors = ['#FFFFFF','#E0F3FC','#CFECF9','#BCE5F7','#ABDAF4','#9FD0F0', '#9AC7E7','#96C3E2','#92BEDC','#8FB9D7','#8BB4D1','#87B0CC','#84ABC6','#80A6C1','#7CA1BB','#789CB5','#7497AF','#7092AA','#6C8DA4','#698AA0','#65859B','#5C798E','#587488','#557185','#506B7D','#4C6678','#486172','#445C6D','#405767','#3B5161']

    def __init__(self,ax13,x,y,s,id,t,boxcolor):
#adding box color
        self.x = x
        self.y = y
        self.size = s
        self.id = id
        self.boxcolor=boxcolor #adding color
        self.text=ax13.text(x, y, t, style='italic',horizontalalignment='center',verticalalignment='center',size=s, color='k',transform=ax13.transAxes,bbox={'facecolor':self.colors[0], 'pad':10})
        # self.text.set_bbox(dict(facecolor='yellow',alpha=0.2,edgecolor='red'))
        self.text.set_bbox(dict(facecolor=boxcolor,alpha=0.2,edgecolor='black'))
        self.list_box.append(self)

    def setXY(self, x,y):
        self.x=x
        self.y=y

    def setSize(self, s):
        self.size=s
        self.text.set_fontsize(s)

class ArrowObject:
    ''' @f from which box
        @t to which box
    '''
    visible_arrow=[]
    def __init__(self,ax13,f,t,id):
        self.from_box = f
        self.to_box = t
        self.id = id
        # start=(data.bxy[f][0],data.bxy[f][1])
        # end=(data.bxy[t][0],data.bxy[t][1])
        end=(data.bxy[f][0],data.bxy[f][1])
        start=(data.bxy[t][0],data.bxy[t][1])
        amin, amax=np.amin(abs(data.a)), np.amax(abs(data.a))
        opacity=(abs(data.a[f][t])-amin)/(amax-amin)*.8 +.2
        # width=opacity*24/data.numc #scale ARROW WIDTH inverse #VARIABLES
        width=opacity*(24.+(2./3.)*(data.numc-6.))/data.numc # a BETTER scale?
        if data.a[f][t]<0:
            arrow_color='r'
        else:
            arrow_color='#00ccff'
        # if data.a[f][t]==0:
        #     muts=.1
        # else:
        #     muts=data.a[f][t]
        self.arrow=patches.FancyArrowPatch(
            start,
            end,
            connectionstyle='arc3, rad=0.1',
            color = arrow_color,
            # arrowstyle='fancy',
            # mutation_scale=data.a[f][t],
            # mutation_scale=1,
            shrinkB=10,
            shrinkA=10,
            # capstyle='projecting',
            alpha=opacity,
            linewidth=width)
        self.arrow.set_arrowstyle('fancy',head_length=6*width, head_width=6*width)
        if data.a[f][t]!=0:
            self.visible_arrow.append(self)
        ax13.add_patch(self.arrow)

class App:
    #constructor
    def __init__(self):
        self.data=data
        self.fewarrows=0
        # self.recalculate(data)
        self.MakeWindow()
        self.refreshDataFrame()
        self.refreshPicFrame()
        self.fixent=1 #UGLY FIX FOR ENTRIES/ENTRIESIJ

        self.root.mainloop()

    def MakeWindow (self):
        self.root=tk.Tk()
        self.root.wm_title("Data Input and Graphical Output")
        self.outsideframed1=tk.Frame(self.root,width=300, height=800,bg='red')
        # self.outsideframed1=tk.Frame(self.root,width=300, height=1000,bg='red')
        # self.outsideframed1=tk.Frame(self.root,width=300, height=800,bg='blue')
        # self.outsideframepic=tk.Frame(self.root,width=675, height=800,bg='bisque')
        self.outsideframepic=tk.Frame(self.root,width=675, height=800)

        # self.outsideframepic=tk.Frame(self.root,width=1000, height=1000,bg='bisque')
        self.outsideframed1.pack(side=tk.LEFT,fill=None,expand=False)
        self.outsideframepic.pack(side=tk.LEFT,fill=None,expand=False)
        self.outsideframed1.pack_propagate(False)
        self.outsideframepic.pack_propagate(False)
        self.framed1=tk.Frame(self.outsideframed1,width=200, height=100,bg='red')
        self.framed1.pack(side=tk.LEFT,fill=None,expand=False)
        self.framepic=tk.Frame(self.outsideframepic,borderwidth=5,relief=tk.RIDGE)
        # self.framepic.pack(side=tk.RIGHT)

#FUDGES TO MAKE WORK ON DESKTOP-1 of 2----------------------------------------
#         self.framepic.pack(anchor=tk.CENTER,pady=150) #CORRECT
#         self.framepic.pack(side=tk.TOP) #doesnt expand canvas
        self.framepic.pack(side=tk.TOP,fill=tk.BOTH,expand=1) #BIG-BAD
#         self.framepic.pack(fill=tk.BOTH) #BIG? NO still small
#         self.framepic.pack(anchor=tk.CENTER,fill=tk.BOTH) #still small
#FUDGES TO MAKE WORK ON DESKTOP- 1 of 2----------------------------------------

        self.refreshDataFrame()
        self.refreshPicFrame()


    def createBoxGraph(self):
        #These next three lines set up the actual graph
        # self.fig13 = plt.figure(facecolor = 'white')
        # ax13 = self.fig13.add_subplot(111, aspect='equal')
        # ax13.axis('off')
        #
        # plt.figure() #THIS DOESN'T SEEM TO BE NEEDED SO ONLY 1 EXTRA FIGURE
        #
        # f = Figure(figsize=(5, 4), dpi=100)
        TextBox.list_box=[]  #CLEAR ALL PREVIOUS!!!
        f = plt.figure(facecolor = 'white')
        f.set_size_inches(8,10)
        a = f.add_subplot(111)
        a.axis('off')
# #ADDED SCALING-------------------------------------
#         vector=data.z[-1]
#         data.b=App.scalebox(vector)
# #ADDED SCALING-------------------------------------
        for index in range(len(data.b)):
            xy=data.bxy[index]
            TextBox(a,xy[0],xy[1],data.b[index],index,data.labels[index],data.boxcolor[index])
        id=0
        if (self.fewarrows==0):
            for i in range(len(data.b)):
                for j in range(len(data.b)):
                    if i!=j and data.a[i][j]!=0:
                        arrow=ArrowObject(a,i,j,id)
                        id=id+1
        else:
            i=self.box_id
            for j in range(len(data.b)):
                if i!=j and data.a[i][j]!=0:
                    arrow=ArrowObject(a,i,j,id)
                    id=id+1
            j=self.box_id
            for i in range(len(data.b)):
                if i!=j and data.a[i][j]!=0:
                    arrow=ArrowObject(a,i,j,id)
                    id=id+1
#*************PUT SET COLOR HERE!!!!!!
                    # id=id+1
        plt.show(block=False)
        # plt.close(1)
        # plt.close(2)

#THIS IS PRETTY IFFY, BOTH GET THE LAST TWO FIGURE NUMBERS AND CLOSE THOSE----
        openfigs=plt.get_fignums()
        last=openfigs[-1]
        plt.close(last)
#THIS IS PRETTY IFFY, BOTH GET THE LAST TWO FIGURE NUMBERS AND CLOSE THOSE----
        return f
    #End of Box Graph

 #ADDED SCALING B-------------------------------------------
    def scalebox(vector):
        data2=[0 for i in range(len(vector))]
        minbox,maxbox=2,30
        minb,maxb=min(vector),max(vector)
        if minb!=maxb:
            data2=[(vector[i]-minb)/(maxb-minb) for i in range(len(vector))]
            vectornew=[(maxbox-minbox)*data2[i]+minbox for i in range(len(vector))]
        else:
            vectornew=[(minbox+maxbox)/2. for i in range(len(vector))]
        return vectornew
#ADDED SCALING B-------------------------------------------


    def recalculate(self,pass_data):
#UGLY FIX FOR ENTRIES/ENTRIESIJ
        print ("Here")
        if self.fixent==1:
            self.data.z[0]=[eval((self.entries[i][1].get())) for i in range(len(self.entries))]
        if self.fixent==2:
            column=[eval((self.entriesIJ[i][1].get())) for i in range(len(self.entriesIJ))]
            self.data.ca[:,self.box_id]=column
        # self.data.z[0]=[eval((self.entries[i][1].get())) for i in range(len(self.entries))]
#UGLY FIX FOR ENTRIES/ENTRIESIJ
        self.fewarrows=0
        pass_data.tt=0
        for i in range (1,pass_data.numdata):
            mtanh=np.tanh(pass_data.z[i-1])
            cterm=np.dot(pass_data.ca,mtanh)
            pass_data.dx=pass_data.dt*(pass_data.ma*pass_data.z[i-1] + pass_data.ba + cterm)
            pass_data.tt=pass_data.tt+pass_data.dt
            pass_data.t[i]=pass_data.tt
            pass_data.z[i]=pass_data.z[i-1]+pass_data.dx
            for j in range(pass_data.numc):
                pass_data.z[i][j]=max(pass_data.z[i][j],0.)
#make new plot
        App.MakePlot(data)

#scale b's from z[-1]
        vector=data.z[-1]
        data.b=App.scalebox(vector)
#set z[0]=z[-1] for the NEXT iteration
        pass_data.z[0]=pass_data.z[-1]
#CLEAR and REFRESH DATA and PIC frames
        App.ClearFrame(self.framed1)
        App.ClearFrame(self.framepic)
        self.refreshDataFrame()
        self.refreshPicFrame()



    def MakePlot(pass_data):
   #     plt.ion() THIS MAKES MATPLOTLIB INTERACTIVE
        print('\nYour plot is ready')
        localtime = time.asctime( time.localtime(time.time()) )
        x_start=pass_data.z[0]
        x_final=pass_data.z[-1]
        plt.figure()
        plt.axes([0.1,.075,.8,.7])
        plt.plot(pass_data.t,pass_data.z[:,0:pass_data.numc])
        #print labels on lines
        xtext=25
        for i in range (pass_data.numc):
            ytext=pass_data.z[-1,i]
            varis=str(i+1)
            plt.text(xtext,ytext,varis)
            xtext=xtext-1
        programname='ORANGE-5 FOLDER   '+localtime
        param1='\n   input files= '+str(pass_data.fnamec)+'    '    +str(pass_data.fnameb)+'    '+str(pass_data.fnamem) +'    '+str(pass_data.fnamebtextbxy) + '     dt='+str(pass_data.dt)
        start=App.displayinput(pass_data.z[0],75)
        finish=App.displayinput(pass_data.z[-1],75)
        param2='\nstart=  ' + start + '\nfinish=  ' + finish
        titlelsl=programname+param1 + param2
        plt.title(titlelsl, fontsize=8)
#LOOK THIS IS REALLY IMPORTANT
#NEED A PLT.SHOW() TO DISPLAY GRAPH
#BUT THAT *STOPS* EXECUTION
#THIS SHOWS THE GRAPH AND THEN CONTINUES SCRIPT
        plt.show(block=False)

    def displayinput(vector1,number):
        #creates string to print from np.array(vector1)
        #that is approximately number characters per line
        c=''
        v1=vector1.tolist()
        v2=[round(v1[i],6) for i in range (len(v1))]
        a=str(v2)
        a2=a.replace(' ','')
        a3=list(a2)
        a3[0],a3[-1]='',''
        numend=0
        for i in range(0,len(a3)):
            if (a3[i]==',' and numend >= number):
                numend=0
                a3[i]=',\n'
            numend=numend+1
        c=''.join(a3)
        c2=c.replace(',',',  ')
        return c2

    # def fetch(entries):
    #     # print('\n')
    #     for entry in entries:
    #         field = entry[0]
    #         text  = entry[1].get()
    #         # print('%s: "%s"' % (field, text))

#CLEARING AND REFRESHING ONLY THE DATA FRAME
    def refreshDataFrame(self):
        self.fixent=1 #UGLY FIX FOR ENTRIES/ENTRIESIJ
        App.ClearFrame(self.framed1)

#Adding new buttons at top
        newframe=tk.Frame(self.framed1,bg='red')
        newframe.pack(side=tk.TOP,pady=5)
        # tk.Label(newframe,text='initial conditions').pack(side=tk.LEFT,padx=20)
        tk.Label(newframe,text='initial conditions').pack(side=tk.LEFT,padx=5)
        tk.Button(newframe,text='original',command= self.resetIC).pack(side=tk.LEFT,padx=30)
        # tk.Button(newframe,text='ALL Cij',command= self.FullrefreshPicFrame).pack(side=tk.LEFT,padx=5)
        # tk.Button(newframe,text=' Change Cij',command= self.refreshCIJFrame).pack(side=tk.LEFT,padx=20)
#Adding new buttons at top


        # self.outsideframed1.pack_propagate(False)
        fields=data.labels
        default=[str(i) for i in range(len(data.labels))]
        entries = []
        # self.framed1=tk.Frame(self.outsideframed1,width=200, height=100,bg='red')
        # self.framed1.pack(side=tk.LEFT,fill=None,expand=False)
    # OK trying to ROUND off in dataframe
        self.data.zround=[str(round(self.data.z[-1,i],6)) for i in range(len(self.data.z[0]))]
        for field in fields:
            row = tk.Frame(self.framed1)
            lab = tk.Label(row, width=15, text=field, anchor='w')
            ent = tk.Entry(row)
            # row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
            # row.pack(side=tk.TOP, padx=5, pady=1, expand=1)
            row.pack(side=tk.TOP, padx=5, pady=0, expand=1)
            lab.pack(side=tk.LEFT,expand=1)
            ent.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.Y)
            ent.insert(10,self.data.zround[fields.index(field)])
            # ent.insert(10,default[fields.index(field)])
            entries.append((field, ent))
            self.entries=entries
        #TRANSFORM ALL ENTRIES INTO STARTING VALUES FOR COMPUTATION
        self.data.z[0]=[eval((entries[i][1].get())) for i in range(len(entries))]
        #BUTTONS
        cal1=tk.Button(self.framed1,text='CALCULATE',command=(lambda: self.recalculate(data)))
        cal1.pack(side=tk.LEFT,pady=20,padx=20)
        tk.Button(self.framed1,text='ENTER',command=self.refreshPicFrame).pack(side=tk.RIGHT,padx=20)
        # tk.Button(self.framed1,text='do not click on this buttion',command=self.myquit).pack(side=tk.RIGHT,padx=20)
        self.outsideframed1.pack(expand=1)
        return

    def refreshPicFrame(self):
 # #get the data from the current dataframe

#UGLY FIX FOR ENTRIES/ENTRIESIJ8
        if self.fixent==1:
            self.data.z[0]=[eval((self.entries[i][1].get())) for i in range(len(self.entries))]
        if self.fixent==2:
            column=[eval((self.entriesIJ[i][1].get())) for i in range(len(self.entriesIJ))]
            self.data.ca[:,self.box_id]=column
#UGLY FIX FOR ENTRIES/ENTRIESIJ

#scale b's from z[0] - NOT Z[-1] like in A NEW CALCULATION
        vector=data.z[0]
        self.data.b=App.scalebox(vector)
# #set z[0]=z[-1] for the NEXT iteration
#         pass_data.z[0]=pass_data.z[-1]
        # self.framepic.pack(side=tk.RIGHT)
        App.ClearFrame(self.framepic)



# #TURNING THIS OFF, trying to pick up from IV_3
# #OK, adding a new frame and new buttons
#         newbuttons=tk.Frame(self.framepic)
#         # newbuttons.pack(anchor='w',side=tk.TOP,pady=5)
#         newbuttons.pack(side=tk.TOP,pady=5)
#         tk.Button(newbuttons,text='ALL Cij',command=quit).pack(side=tk.TOP,padx=20)
#         tk.Button(newbuttons,text='ALL Cij',command=quit).pack(side=tk.TOP,padx=20)
#         tk.Button(newbuttons,text='ALL Cij',command=quit).pack(side=tk.TOP,padx=20)
# #TURNING THIS OFF, trying to pick up from IV_3


# #OK, HIDING THEN SHOWING AFTER PICTURE
#         framehide=tk.Frame(self.framepic)
#         framehide.pack(side=tk.TOP,anchor='w',pady=5)


        # self.outsideframepic.pack_propagate(False)
#         self.canvas=tk.Canvas(self.framepic,width=200, height=100) #CORRECT
        self.canvas=tk.Canvas(self.framepic,width=800, height=2400)
        f=self.createBoxGraph()
        self.canvas = FigureCanvasTkAgg(f, master=self.framepic)
        self.canvas.show()
        # framehide.pack_forget()

#FUDGES TO MAKE WORK ON DESKTOP-2 of 2----------------------------------------
#         self.canvas._tkcanvas.pack(side=tk.TOP, fill=tk.BOTH, expand=1) # CORRECT
#         self.canvas._tkcanvas.pack(fill=tk.BOTH, expand=1) #BIG-BAD
        self.canvas._tkcanvas.pack() #BIG ??
#HERE NOT NEEDED!!
#FUDGES TO MAKE WORK ON DESKTOP-2 of 2----------------------------------------

        cid=f.canvas.mpl_connect('button_press_event',self.onclick)
    # def UpdateButton(self):

    def refreshCIJFrame(self):
#VARIABLE NUMBER FROM CLICK
#NB first [] in CA is BOX INTO
#NB second [] in CA is BOX FROM
        self.fixent=2 #UGLY FIX FOR ENTRIES/ENTRIESIJ
        # self.box_id=4
        App.ClearFrame(self.framed1)

        fromto='From    '+data.labels[self.box_id]+'    to'

        newframe=tk.Frame(self.framed1,bg='red')
        newframe.pack(side=tk.TOP,pady=5)
        tk.Label(newframe,text=fromto,bg='thistle1').pack(side=tk.LEFT,padx=5)
        tk.Button(newframe,text='ALL Cij',command= self.FullrefreshPicFrame).pack(side=tk.LEFT,padx=5)
        tk.Button(newframe,text='IC',command= self.refreshDataFrame).pack(side=tk.LEFT,padx=5)



         #
         #tk.Label(self.framed1,text=fromto,bg='thistle1').pack(anchor=tk.N,side=tk.LEFT, fill=tk.X, pady=5)
        # # self.outsideframed1.pack_propagate(False)
        # tk.Button(self.framed1,text='HERE?',command=self.refreshPicFrame).pack(anchor=tk.N,side=tk.RIGHT,fill=tk.X)

        fields=self.data.labels
        # default=[str(i) for i in range(len(self.data.labels))]
        entriesIJ = []
        for field in fields:
            # print('\n\nfield,fields',field,fields)
            row = tk.Frame(self.framed1)
            lab = tk.Label(row, width=15, text=field, anchor='w',bg='thistle1')
            entIJ = tk.Entry(row,bg='thistle1')
            # row.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
            # row.pack(side=tk.TOP, padx=5, pady=1, expand=1)
            row.pack(side=tk.TOP, padx=5, pady=1, expand=1)
            lab.pack(side=tk.LEFT,expand=1)
            entIJ.pack(side=tk.RIGHT, expand=tk.YES, fill=tk.Y)
            # ent.insert(10,self.data.ca[index(field)][fields.index(field)])
            # entIJ.insert(10,self.data.ca[fields.index(field)][self.box_id])
            # entIJ.insert(10,self.data.ca[self.box_id][fields.index(field)])
            entIJ.insert(10,self.data.ca[fields.index(field)][self.box_id])
            # ent.insert(10,default[fields.index(field)])
            entriesIJ.append((field, entIJ))
            self.entriesIJ=entriesIJ
#TRANSFORM ALL ENTRIES INTO STARTING VALUES FOR COMPUTATION
        column=[eval((self.entriesIJ[i][1].get())) for i in range(len(self.entriesIJ))]
        self.data.ca[:,self.box_id]=column
#        BUTTONS
        cal2=tk.Button(self.framed1,text='CALCULATE',bg='thistle1',command=(lambda: self.recalculate(data)))
        cal2.pack(side=tk.LEFT,padx=20, pady=5)
        # cal2.pack(side=tk.LEFT,pady=20,padx=20)
        # tk.Button(self.framed1,text='UPDATE Cij',bg='thistle1',command=self.refreshPicFrame).pack(side=tk.RIGHT,padx=20)
        tk.Button(self.framed1,text='ENTER',bg='thistle1',command=self.refreshPicFrame).pack(side=tk.RIGHT, padx=20)
        # tk.Button(self.framed1,text='do not click on this buttion',command=self.myquit).pack(side=tk.RIGHT,padx=20)
        self.outsideframed1.pack(expand=1)
        return

    # def fetch(entries):
    #     # print('\n')
    #     for entry in entries:
    #         field = entry[0]
    #         text  = entry[1].get()
    #         # print('%s: "%s"' % (field, text))

    def onclick(self,event):
        for box in TextBox.list_box:
            contains, attrd = box.text.contains(event)
            if(contains):
                id=box.id
                print('\nid,bname(id)=  ',id, data.labels[id])
                # print('box_%d'  % id)
                # print('box_' + data.bname[id])
                # print('show vars ')
                # self.update_dataFrame(id)
                self.box_id=id
                self.fewarrows=1
                self.refreshCIJFrame()
                # TextBox.selected_box_id=id
                return;

    def resetIC(self):
        self.data.z[-1]=[self.data.ica[i] for i in range(len(self.data.z[0]))]
        self.refreshDataFrame()

    def FullrefreshPicFrame(self):
        self.fewarrows=0
        self.refreshPicFrame()

    def myquit(self):
        print ('\n I did press CLOSE!')
        self.root.destroy()

    def ClearFrame(frame):
        for widget in frame.winfo_children():
            widget.destroy()
        # frame.pack_forget()
