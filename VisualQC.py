# -*- coding: utf-8 -*-
'''
https://github.com/ragardner/tksheet/wiki#24-example-loading-data-from-excel
https://github.com/ragardner/tksheet

#To manipulate the sheet
https://gitlab.physics.ox.ac.uk/metodiev/gui_nek/-/tree/master/tksheet
'''
from tksheet import Sheet
import tkinter as tk
from tkinter import Tk, Label, Button, StringVar,OptionMenu,E,W, Scale,DoubleVar,HORIZONTAL,Radiobutton,Checkbutton,IntVar,Spinbox,Entry,END
from tkinter import scrolledtext
from tkinter import ttk
import pandas as pd
import os
import numpy as np
import time
#import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile
#from matplotlib.widgets import RectangleSelector
import matplotlib.pyplot as plt
from matplotlib import style

from ttkthemes import ThemedTk
from tkinter import messagebox
import uuid
import time




n=[]
colNum=0
checked=0
rowCoord=0
MyQC="0"
flagNewCol=0


NUMBERS_ARRAY = []

for n in range(150):
    NUMBERS_ARRAY.append(n)


TMP_LETTERS_ARRAY = [
    "A",
    "B",
    "C",
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
    "Q",
    "R",
    "S",
    "T",
    "U",
    "V",
    "W",
    "X",
    "Y",
    "Z",
]

LETTERS_ARRAY = TMP_LETTERS_ARRAY

howmanynumbers = len(NUMBERS_ARRAY)-1
if howmanynumbers > 25:
    letters_cicles=int(float(howmanynumbers)/25)
    for c in range(letters_cicles):
        for n in range(26):
            LETTERS_ARRAY.append(TMP_LETTERS_ARRAY[c]+TMP_LETTERS_ARRAY[n])
LETTERS_ARRAY.insert(0, " ")


def esegui():
    
    action=hiddenLabel["text"]
    #print(action)
    
    actualdirname = os.getcwd()
    name = askopenfilename(initialdir=actualdirname,
                                filetypes =(("XLSX", "*.xlsx"),("XLS", "*.xls"),("CSV","*.csv"),("All Files","*.*")),
                                title = "Choose a file."
                                )
    
    if name.endswith('.csv'):
        read_file = pd.read_csv(name)
        idrnd = uuid.uuid4()
        tempExcel=str(time.strftime("%Y%m%d%H%M%S")+'-'+str(idrnd)+'.xlsx')
        read_file.to_excel(tempExcel, index=False, header=True)
        sheet = Sheet(LabelFrameXls,
                                   data = pd.read_excel(tempExcel,      # filepath here
                                                        #sheet_name = "sheet1", # optional sheet name here
                                                        engine = "openpyxl",
                                                        header = None).values.tolist())
    else:
        #df = pd.read_csv(name)
        sheet = Sheet(LabelFrameXls,
                                   data = pd.read_excel(name,      # filepath here
                                                        #sheet_name = "sheet1", # optional sheet name here
                                                        engine = "openpyxl",
                                                        header = None).values.tolist())
    
    sheet.enable_bindings("all")
    

    
    
    def searchOutliers():
        tmpMyColOutliers=int(LETTERS_ARRAY.index(enOutliersCol.get()))-1
        #print(str(tmpMyColOutliers))
        countRowOutliers=0
        columnNameOutliers=''
        listOutliers=[]
        try:
            for value in sheet.get_column_data(tmpMyColOutliers):
                #print(str(value))
                if countRowOutliers==0:
                    columnNameOutliers=value
                    
                if countRowOutliers>=1:
                    listOutliers.append(float(value))
                    
                countRowOutliers+=1
            #print(listOutliers)
            dfOutliers = pd.DataFrame(listOutliers, columns=[columnNameOutliers])
            min_threshold,max_threshold = dfOutliers[columnNameOutliers].quantile([float(valueMinThreshold.get()),float(valueMaxThreshold.get())])
            InfoOutliers.insert(END, '\n')
            InfoOutliers.insert(END, '\nMIN threshold'+str(min_threshold))
            InfoOutliers.insert(END, '\nMAX threshold'+str(max_threshold))
            #print(str(min_threshold)+' - '+str(max_threshold))
            outliersFounded=dfOutliers[(dfOutliers[columnNameOutliers]<min_threshold)|(dfOutliers[columnNameOutliers]>max_threshold)]
            InfoOutliers.insert(END, '\n')
            InfoOutliers.insert(END, '\nOutliers founded')
            InfoOutliers.insert(END, '\n'+str(outliersFounded))
            #print(outliersFounded)
            InfoOutliers.insert(END, '\n')
            InfoOutliers.insert(END, '\nOutliers dataframe indexes')
            InfoOutliers.insert(END, '\n'+str(outliersFounded.index))
            #print(outliersFounded.index)
        except Exception as e:
            #print("An exception occurred")
            #print(e)
            messagebox.showwarning("showwarning", "Warning: "+e)
            
    
    def removeOutliers():
        tmpMyColOutliers=int(LETTERS_ARRAY.index(enOutliersCol.get()))-1
        #print(str(tmpMyColOutliers))
        countRowOutliers=0
        columnNameOutliers=''
        listOutliers=[]
        try:
            for value in sheet.get_column_data(tmpMyColOutliers):
                #print(str(value))
                if countRowOutliers==0:
                    columnNameOutliers=value
                    
                if countRowOutliers>=1:
                    listOutliers.append(float(value))
                    
                countRowOutliers+=1
            #print(listOutliers)
            dfOutliers = pd.DataFrame(listOutliers, columns=[columnNameOutliers])
            min_threshold,max_threshold = dfOutliers[columnNameOutliers].quantile([float(valueMinThreshold.get()),float(valueMaxThreshold.get())])
            InfoOutliers.insert(END, '\n')
            InfoOutliers.insert(END, '\nMIN threshold'+str(min_threshold))
            InfoOutliers.insert(END, '\nMAX threshold'+str(max_threshold))
            #print(str(min_threshold)+' - '+str(max_threshold))
            outliersFounded=dfOutliers[(dfOutliers[columnNameOutliers]<min_threshold)|(dfOutliers[columnNameOutliers]>max_threshold)]
            InfoOutliers.insert(END, '\n')
            InfoOutliers.insert(END, '\nOutliers founded')
            InfoOutliers.insert(END, '\n'+str(outliersFounded))
            #print(outliersFounded)
            InfoOutliers.insert(END, '\n')
            for tmpindex in outliersFounded.index:
                delRow=int(tmpindex)+1
                #print('Delete outlier at row '+str(delRow))
                InfoOutliers.insert(END, '\nDeleted outlier at row '+str(delRow))
                sheet.set_cell_data(delRow,tmpMyColOutliers, '')
        except Exception as e:
            #print("An exception occurred")
            #print(e)
            messagebox.showwarning("showwarning", "Warning: "+e)
        
    def rc():
        try:
            
            #XLSoutputName = str(time.strftime("%Y%m%d%H%M%S")+'.xlsx')
            files = [('Excel Document', '*.xlsx')] 
            file = asksaveasfile(filetypes = files, defaultextension = files)
            df = pd.DataFrame(sheet.get_sheet_data(return_copy = True, get_header = True, get_index = False))
            df.to_excel(file.name,sheet_name='Sheet_name_1')
            
        except:
            print("An exception occurred")

            
    def newcol():
        tmpMyCol=int(LETTERS_ARRAY.index(enQCCol.get()))-1
        tmpMyColQC=int(LETTERS_ARRAY.index(enQCColCheck.get()))-1
        print("Selected column QC in"+str(tmpMyCol))
        print("Selected column QC out"+str(tmpMyColQC))
        print("MIN Value = " + str(l1.get()))
        print("MAX Value = " + str(l2.get()))
        print("Spike Value = " + str(l1Spike.get()))
        
        val1=0
        val2=0
        val3=0
        
        if tmpMyCol==tmpMyColQC:
            print("Warning: column to check and QC's column can't be the same")
            messagebox.showwarning("showwarning", "Warning: column to check and QC's column can't be the same")
        if tmpMyCol==-1:
            #print("Warning: column to check is not defined")
            messagebox.showwarning("showwarning", "Warning: column to check is not defined")
        if tmpMyColQC==-1:
            #print("Warning: QC column is not defined")
            messagebox.showwarning("showwarning", "Warning: QC column is not defined")

            
        try:
            countRow=0
            for value in sheet.get_column_data(tmpMyCol):
                
                
                
                if countRow>=1:
                    try:
                        
                        
                        print(value)
                        valueToCheck=float(value)
                        
                        
                        #QC1 default value or not?
                        if varOkQC1.get()==1:
                            defaultQcVal=1
                        else:
                            defaultQcVal=0
                        
                        #The values for the spike's check
                        if countRow==1:
                            val1=valueToCheck
                            val2=0
                            val3=0
                        if countRow==2:
                            val2=valueToCheck
                            val3=0
                        if countRow==3:
                            val3=valueToCheck
                        if countRow>=4:
                            val1=val2
                            val2=val3
                            val3=valueToCheck
                        
                        #QC4 check
                        if (valueToCheck < float(l1.get()) or valueToCheck > float(l2.get())) and varOkQC4.get()==1:
                            QcValue=4
                            sheet.set_cell_data(countRow,tmpMyColQC, QcValue)
                            sheet.highlight_cells(row = countRow, column = tmpMyColQC, cells = [], canvas = "table", bg = "violet", fg = None, redraw = False, overwrite = True)
                        
                            
                        else:
                            
                            spike=abs(abs((val2-val1)-(val3-val1))-abs(val3-val1))
                            
                            if spike >float(l1Spike.get()) and varOkQC3.get()==1:
                                QcValue=3
                                sheet.set_cell_data(countRow,tmpMyColQC, QcValue)
                                sheet.highlight_cells(row = countRow, column = tmpMyColQC, cells = [], canvas = "table", bg = "red", fg = None, redraw = False, overwrite = True)
                        
                            else:
                                QcValue=1
                                sheet.set_cell_data(countRow,tmpMyColQC, defaultQcVal)
                                sheet.highlight_cells(row = countRow, column = tmpMyColQC, cells = [], canvas = "table", bg = "green", fg = None, redraw = False, overwrite = True)
                        
                        
                        
                    except:
                        if value=="" and varOkQC9.get()==1:
                            sheet.set_cell_data(countRow,tmpMyColQC, 9)
                        else:
                            sheet.set_cell_data(countRow,tmpMyColQC, defaultQcVal)
                countRow+=1
        except Exception as e:
            #print("An exception occurred ")
            #print(e)
            messagebox.showwarning("showwarning", "Warning: "+e)
            
            
    sheet.popup_menu_add_command("Save all as XLSX", rc, table_menu = True, index_menu = True, header_menu = True)
    sheet.popup_menu_add_command("Use Automatic QC assignement", newcol, table_menu = True, index_menu = True, header_menu = True)
    sheet.popup_menu_add_command("Check Outliers", searchOutliers, table_menu = True, index_menu = True, header_menu = True)
    sheet.popup_menu_add_command("Remove Outliers", removeOutliers, table_menu = True, index_menu = True, header_menu = True)
    sheet.grid(row = 0, column = 0, sticky = "nswe")
    sheet.refresh(redraw_header = True, redraw_row_index = True)
    
    
    
    
    

    
    def pr(event):
        #row = sheet.identify_row(event, exclude_index = False, allow_end = True)
        column = sheet.identify_column(event, exclude_header = False, allow_end = True)
        col=sheet.get_selected_columns()
        for i in col:
            colNum=i
            print(colNum)
            
        noLabel=0
        tmpLabel=''
        n=[]
        try:
            for e in sheet.get_column_data(colNum):
                
                if noLabel==1:
                    #print(e)
                    n.append(e)
                    
                if noLabel==0:
                    tmpLabel=str(e)
                    
                noLabel=1
                
            checked=0
        except Exception as e:
            #print("An exception occurred")
            #print(e)
            messagebox.showwarning("showwarning", "Warning: "+e)
            checked=1
        
        
        def on_pick(event):
            
            checkQCColor=MyQC.get()
            
            artist = event.artist
            xmouse, ymouse = event.mouseevent.xdata, event.mouseevent.ydata
            x, y = artist.get_xdata(), artist.get_ydata()
            ind = event.ind

            ax.plot(x[ind[0]], y[ind[0]], 'b*')
            ax.annotate(f"QC "+checkQCColor, (x[ind[0]], y[ind[0]]),color='Black')
            event.canvas.draw()
            
                
                
            sheet.set_cell_data(int(x[ind[0]]+1),colNum+1, checkQCColor)
            
            
            #print(checkQCColor)
            if checkQCColor=="0":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "white", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="1":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "green", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="2":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "blue", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="3":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "red", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="4":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "violet", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="5":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "white", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="6":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "white", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="7":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "white", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="8":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "white", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="9":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "white", fg = None, redraw = False, overwrite = True)
            if checkQCColor=="Q":
                sheet.highlight_cells(row = int(x[ind[0]]+1), column = colNum+1, cells = [], canvas = "table", bg = "blue", fg = None, redraw = False, overwrite = True)

                
        if checked==0:
            #DarkestorNot=varShowDarkest.get()
            #if DarkestorNot == 1:
            #    MyThemeTMP=MyTheme.get()
            #    plt.style.use(MyThemeTMP)
            #else:
            #    plt.style.use('default')
            
            MyThemeTMP=MyTheme.get()
            plt.style.use(MyThemeTMP)
            MyLineTMP=MyLine.get()
            
            fig, ax = plt.subplots()
            ax.set_title(tmpLabel)
            tolerance = 10 # points
            ax.plot(n, '-o', picker=tolerance, ms=6, lw=2, alpha=0.7, mfc='orange',linestyle = MyLineTMP)
            
                
            QCorNot=varShowQC.get()
            if QCorNot == 1:
                for index in range(len(n)):
                    showQCval="QC "+str(sheet.get_cell_data(index+1, colNum+1))

                    qccolor=""
                    if showQCval=="QC 1":
                        qccolor="green"
                    elif showQCval=="QC 2":
                        qccolor="blue"
                    elif showQCval=="QC 3":
                        qccolor="red"
                    elif showQCval=="QC 4":
                        qccolor="violet"
                    elif showQCval=="QC 5":
                        qccolor="black"
                    elif showQCval=="QC 6":
                        qccolor="black"
                    elif showQCval=="QC 7":
                        qccolor="black"
                    elif showQCval=="QC 8":
                        qccolor="black"
                    elif showQCval=="QC 9":
                        qccolor="black"
                    elif showQCval=="QC Q":
                        qccolor="blue"
                    else:
                        qccolor="black"
                    
                        
                    
                    ax.text(index, n[index], showQCval, size=12,color =qccolor, fontweight ='bold', clip_on=True)
                    

                
            fig.canvas.callbacks.connect('pick_event', on_pick)

            GRIDorNot=varShowGrid.get()
            if GRIDorNot == 1:
                plt.grid()
            
            
            plt.show()

    hiddenLabel.config(text = "load")
    sheet.bind("<ButtonPress-1>", pr)
    
    
    

    
def printnvar():
    try:
        
        df = pd.DataFrame(sheet.get_column_data(0))
    except Exception as e:
        #print("An exception occurred")
        #print(e)
        messagebox.showwarning("showwarning", "Warning: "+e)

    print(df)


    
root = ThemedTk(theme="radiance")
root.title('Visual QC')
root.option_add('*Font', 'Verdana 8')
root.geometry('770x530')

frameFont = ttk.Style()
frameFont.configure('new.TFrame', family='Verdana', size=8, weight='bold', underline=1)

nb = ttk.Notebook(root)
nb.grid(row=15, column=3, sticky='NESW')


LabelFrameInfo = ttk.Frame(nb, style='new.TFrame')
nb.add(LabelFrameInfo, text='Main Area')

LabelFrameXls = ttk.Frame(nb, style='new.TFrame')
nb.add(LabelFrameXls, text='Visual QC')

LabelFrameAutoQC = ttk.Frame(nb, style='new.TFrame')
nb.add(LabelFrameAutoQC, text='Automatic QC')




dataEmpty = {'empty set': [np.nan]}
dataEmptyDF = pd.DataFrame(dataEmpty)
sheet = Sheet(LabelFrameXls,data=dataEmpty)
sheet.enable_bindings("all")
sheet.grid(row = 0, column = 0, sticky = "nswe")
sheet.refresh(redraw_header = True, redraw_row_index = True)


InfoOutliers = scrolledtext.ScrolledText(LabelFrameXls, height=20, width=45)
InfoOutliers.grid(row=0, column=1, columnspan=4, sticky=W)

SpacelQCInfo = Label(LabelFrameInfo, text =" ")
SpacelQCInfo.grid(row=0, column=0, sticky=W) 

openInputFileButton = Button(LabelFrameInfo, text="Select file (XLSX, XLS, CSV)",bg = "moccasin", command=(esegui))
openInputFileButton.grid(row=1, column=1, sticky=W)

SpacelQCInfoBis = Label(LabelFrameInfo, text =" ")
SpacelQCInfoBis.grid(row=2, column=0, sticky=W)

mytext=''
labelINFO = Label(LabelFrameInfo, text=mytext, bg="sienna3", justify="left", fg="white", height=20, width=40)
labelINFO['text'] = 'QC VALUES:\n\n*QC 0 - NO QUALITY CONTROL \n*QC 1 - GOOD VALUE \n*QC 2 - PROBABLY GOOD VALUE \n*QC 3 - PROBABLY BAD VALUE \n*QC 4 - BAD VALUE \n*QC 5 - CHANGED VALUE \n*QC 6 - VALUE BELOW DETECTION \n*QC 7 - VALUE IN EXCESS \n*QC 8 - INTERPOLATED VALUE \n*QC 9 - MISSING VALUE \n*QC A - PHENOMENON UNCERTAIN \n*QC Q - UNDER DETECTION VALUE \n'
labelINFO.grid(row=3, column=1, columnspan=5, rowspan=7)

mytextB=''
labelINFOB = Label(LabelFrameInfo, text=mytextB, justify="left", fg="white", height=20, width=5)
labelINFOB.grid(row=3, column=6, rowspan=7)

InfoMain = scrolledtext.ScrolledText(LabelFrameInfo, height=20, width=55)
InfoMain.grid(row=3, column=7, columnspan=4, rowspan=7, sticky=W)
InfoMain.insert(END, '\n ------------------------------------ ')
InfoMain.insert(END, '\n Welcome to VisualQC software')
InfoMain.insert(END, '\n ------------------------------------ ')
InfoMain.insert(END, '\n For any suggestion or bug, contact me:')
InfoMain.insert(END, '\n pythonopenprojects@gmail.com')
InfoMain.insert(END, '\n ')
InfoMain.insert(END, '\n ')
InfoMain.insert(END, '\n ')
InfoMain.insert(END, '\n ----------------------------------------')
InfoMain.insert(END, '\n This software is under MIT License ')
InfoMain.insert(END, '\n ----------------------------------------')
InfoMain.insert(END, '\n ')
InfoMain.insert(END, '\n Permission is hereby granted, free of charge, \n to any person obtaining a copy ')
InfoMain.insert(END, '\n of this software and associated documentation \n files (the "Software"), to deal ')
InfoMain.insert(END, '\n in the Software without restriction, including \n without limitation the rights ')
InfoMain.insert(END, '\n to use, copy, modify, merge, publish, distribute, \n sublicense, and/or sell ')
InfoMain.insert(END, '\n copies of the Software, and to permit persons \n to whom the Software is ')
InfoMain.insert(END, '\n furnished to do so, subject to the following \n conditions: ')
InfoMain.insert(END, '\n ')
InfoMain.insert(END, '\n The above copyright notice and this permission \n notice shall be included in all ')
InfoMain.insert(END, '\n copies or substantial portions of the Software. ')
InfoMain.insert(END, '\n ')
InfoMain.insert(END, '\n THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY \n OF ANY KIND, EXPRESS OR ')
InfoMain.insert(END, '\n IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES \n OF MERCHANTABILITY, ')
InfoMain.insert(END, '\n FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \n IN NO EVENT SHALL THE ')
InfoMain.insert(END, '\n AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY \n CLAIM, DAMAGES OR OTHER ')
InfoMain.insert(END, '\n LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT \n OR OTHERWISE, ARISING FROM, ')
InfoMain.insert(END, '\n OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE \n USE OR OTHER DEALINGS IN THE ')
InfoMain.insert(END, '\n SOFTWARE. ')

LabelQC = Label(LabelFrameXls, text ="Select the QC value for the visual check")
LabelQC.grid(row=1, column=0, sticky=W) 


# Dropdown menu options
options = [
    "0",
    "1",
    "2",
    "3",
    "4",
    "5",
    "6",
    "7",
    "8",
    "9",
    "A",
    "Q",
]

MyQC = StringVar()
# initial menu text
MyQC.set( "0" )
# Create Dropdown menu
drop = OptionMenu(LabelFrameXls,MyQC,*options )
drop.grid(row=2, column=0, sticky=W)

varShowQC = IntVar()
buttonShowQC = Checkbutton(LabelFrameXls, text="Show QC on plot", variable=varShowQC)
buttonShowQC.grid(row=3, column=0, sticky=W)

varShowGrid = IntVar()
buttonShowGrid = Checkbutton(LabelFrameXls, text="Show GRID on plot", variable=varShowGrid)
buttonShowGrid.grid(row=4, column=0, sticky=W)

#varShowDarkest = IntVar()
#buttonShowDarkest = Checkbutton(LabelFrameXls, text="Choose the theme and the line style for the plot", variable=varShowDarkest)
buttonShowDarkest = Label(LabelFrameXls, text ="Choose the theme")
buttonShowDarkest.grid(row=5, column=0, sticky=W)

buttonShowLine = Label(LabelFrameXls, text ="Choose the line Style")
buttonShowLine.grid(row=7, column=0, sticky=W)

# Dropdown menu options
optionsTheme = [
    "default",
    "classic",
    "dark_background",
    "Solarize_Light2",
    "fast",
    "fivethirtyeight",
    "bmh",
    "ggplot",
    "grayscale",
]
# datatype of menu text
MyTheme = StringVar()
# initial menu text
MyTheme.set( "default" )
# Create Dropdown menu
dropTheme = OptionMenu(LabelFrameXls,MyTheme,*optionsTheme )
dropTheme.grid(row=6, column=0, sticky=W)




optionsLine = [
    "solid",
    "dotted",
    "dashed",
    "dashdot",
    "None",
]
# datatype of menu text
MyLine = StringVar()
# initial menu text
MyLine.set( "solid" )
# Create Dropdown menu
dropLine = OptionMenu(LabelFrameXls,MyLine,*optionsLine )
dropLine.grid(row=8, column=0, sticky=W)


TitleOutliers = Label(LabelFrameXls, text ="OUTLIERS AREA (quantile method)", bg="khaki", width=45)
TitleOutliers.grid(row=1, column=1, columnspan=2,sticky=W) 

LabelMinThreshold = Label(LabelFrameXls, text ="Select MIN Threshold")
LabelMinThreshold.grid(row=2, column=1, sticky=E)
valueMinThreshold = Entry(LabelFrameXls,)
valueMinThreshold.grid(row=2, column=2, sticky=W)
valueMinThreshold.insert(END, '0.01')

LabelMaxThreshold = Label(LabelFrameXls, text ="Select MAX Threshold")
LabelMaxThreshold.grid(row=3, column=1, sticky=E)
valueMaxThreshold = Entry(LabelFrameXls,)
valueMaxThreshold.grid(row=3, column=2, sticky=W)
valueMaxThreshold.insert(END, '0.99')


LabelOutliersCol = Label(LabelFrameXls, text ="Select the column to check")
LabelOutliersCol.grid(row=4, column=1, sticky=E) 

entriesOutliersColVars = []
tempentriesOutliersColVars = tk.IntVar()
enOutliersCol = Spinbox(LabelFrameXls, values=LETTERS_ARRAY, textvariable=tempentriesOutliersColVars, width=4)
entriesOutliersColVars.append(tempentriesOutliersColVars)
enOutliersCol.grid(row=4, column=2, sticky=W)


SpacelQCAuto = Label(LabelFrameAutoQC, text =" ")
SpacelQCAuto.grid(row=0, column=0, sticky=W) 


hiddenLabel = Label(LabelFrameAutoQC, text ="load")
hiddenLabel.grid(row=1, column=0, sticky=W) 
hiddenLabel.grid_forget()


'''
START HORIZONTAL SCALE VALUE (for ranges)
'''
v1 = DoubleVar()

def show1outMin():  
      
    sel = str(v1.get())
    #l1.config(text = sel, font =("Courier", 8))
    l1.delete(0,END)
    l1.insert(0,sel)

s1 = Scale(LabelFrameAutoQC, variable = v1, 
           from_ = -1000, to = 1000, 
           orient = HORIZONTAL)  
s1.grid(row=0, column=1, sticky=W) 
  
l3 = Label(LabelFrameAutoQC, text = "MIN OUT of RANGE value")
l3.grid(row=1, column=1, sticky=W) 

b1 = Button(LabelFrameAutoQC, text ="Display MIN value", 
            command = show1outMin, 
            bg = "aquamarine")  
b1.grid(row=2, column=1, sticky=W) 


l1 = Entry(LabelFrameAutoQC,)
l1.grid(row=3, column=1, sticky=W) 




v2 = DoubleVar()

def show1outMax():  
      
    sel2 = str(v2.get())
    
    l2.delete(0,END)
    l2.insert(0,sel2)

s2 = Scale(LabelFrameAutoQC, variable = v2, 
           from_ = -1000, to = 1000, 
           orient = HORIZONTAL)  
s2.grid(row=4, column=1, sticky=W) 
  
l4 = Label(LabelFrameAutoQC, text = "MAX OUT of RANGE value")
l4.grid(row=5, column=1, sticky=W) 

b2 = Button(LabelFrameAutoQC, text ="Display MAX value", 
            command = show1outMax, 
            bg = "aquamarine")  
b2.grid(row=6, column=1, sticky=W) 


l2 = Entry(LabelFrameAutoQC,)
l2.grid(row=7, column=1, sticky=W) 


v1Spike = DoubleVar()

def show1spike():  
      
    selSpike = str(v1Spike.get())
    
    l1Spike.delete(0,END)
    l1Spike.insert(0,selSpike)

s1Spike = Scale(LabelFrameAutoQC, variable = v1Spike, 
           from_ = 0, to = 100, 
           orient = HORIZONTAL)  
s1Spike.grid(row=0, column=2, sticky=W) 
  
l3Spike = Label(LabelFrameAutoQC, text = "Spike value")
l3Spike.grid(row=1, column=2, sticky=W) 

b1Spike = Button(LabelFrameAutoQC, text ="Display Spike value", 
            command = show1spike, 
            bg = "plum")  
b1Spike.grid(row=2, column=2, sticky=W) 


l1Spike = Entry(LabelFrameAutoQC,)
l1Spike.grid(row=3, column=2, sticky=W)


varOkQC1 = IntVar()
buttonvarOkQC1 = Checkbutton(LabelFrameAutoQC, text="Set QC 1 as default value", variable=varOkQC1)
buttonvarOkQC1.grid(row=8, column=1,columnspan=10, sticky=W)

varOkQC3 = IntVar()
buttonvarOkQC3 = Checkbutton(LabelFrameAutoQC, text="Check QC 3", variable=varOkQC3)
buttonvarOkQC3.grid(row=9, column=1,columnspan=10, sticky=W)

varOkQC4 = IntVar()
buttonvarOkQC4 = Checkbutton(LabelFrameAutoQC, text="Check QC 4", variable=varOkQC4)
buttonvarOkQC4.grid(row=10, column=1,columnspan=10, sticky=W)

varOkQC9 = IntVar()
buttonvarOkQC9 = Checkbutton(LabelFrameAutoQC, text="Check QC 9", variable=varOkQC9)
buttonvarOkQC9.grid(row=11, column=1,columnspan=10, sticky=W)


    



LabelQCAuto = Label(LabelFrameAutoQC, text ="Select the column for the automatic check")
LabelQCAuto.grid(row=12, column=1, sticky=W) 

entriesQCColVars = []
tempentriesQCColVars = tk.IntVar()
enQCCol = Spinbox(LabelFrameAutoQC, values=LETTERS_ARRAY, textvariable=tempentriesQCColVars, width=4)
entriesQCColVars.append(tempentriesQCColVars)
enQCCol.grid(row=12, column=2,columnspan=10, sticky=W)



LabelQCAutoCheck = Label(LabelFrameAutoQC, text ="Select the column where put the QCs")
LabelQCAutoCheck.grid(row=13, column=1, sticky=W) 

entriesQCColVarsCheck = []
tempentriesQCColVarsCheck = tk.IntVar()
enQCColCheck = Spinbox(LabelFrameAutoQC, values=LETTERS_ARRAY, textvariable=tempentriesQCColVarsCheck, width=4)
entriesQCColVarsCheck.append(tempentriesQCColVarsCheck)
enQCColCheck.grid(row=13, column=2,columnspan=10, sticky=W)





#InfoOutliers = scrolledtext.ScrolledText(LabelFrameOutliers, height=20, width=55)
#InfoOutliers.grid(row=3, column=0, columnspan=4, sticky=W)



    
#searchOutliers = Button(LabelFrameOutliers, text ="Search for Outliers", command = searchOutliers,bg = "plum")  
#earchOutliers.grid(row=4, column=0, sticky=W)

import matplotlib.cm as cm
import matplotlib.font_manager
from matplotlib.patches import Rectangle, PathPatch
from matplotlib.text import TextPath
import matplotlib.transforms as mtrans
MPL_BLUE = '#11557c'


def get_font_properties():
    # The original font is Calibri, if that is not installed, we fall back
    # to Carlito, which is metrically equivalent.
    if 'Calibri' in matplotlib.font_manager.findfont('Calibri:bold'):
        return matplotlib.font_manager.FontProperties(family='Calibri',
                                                      weight='bold')
    if 'Carlito' in matplotlib.font_manager.findfont('Carlito:bold'):
        #print('Original font not found. Falling back to Carlito. '
        #      'The logo text will not be in the correct font.')
        return matplotlib.font_manager.FontProperties(family='Carlito',
                                                      weight='bold')
    #print('Original font not found. '
    #      'The logo text will not be in the correct font.')
    return None

def create_text_axes(fig, height_px):
    """Create an Axes in *fig* that contains 'matplotlib' as Text."""
    ax = fig.add_axes((0, 0, 1, 1))
    ax.set_aspect("equal")
    ax.set_axis_off()

    path = TextPath((0, 0), "Visual Quality Control", size=height_px * 0.8,
                    prop=get_font_properties())

    angle = 4.25  # degrees
    trans = mtrans.Affine2D().skew_deg(angle, 0)

    patch = PathPatch(path, transform=trans + ax.transData, color=MPL_BLUE,
                      lw=0)
    ax.add_patch(patch)
    ax.autoscale()

def splash_screen(height_px, lw_bars, lw_grid, lw_border, rgrid, with_text=False):
    
    dpi = 100
    height = height_px / dpi
    figsize = (5 * height, height) if with_text else (height, height)
    fig = plt.figure(figsize=figsize, dpi=dpi)
    fig.patch.set_alpha(0)

    if with_text:
        create_text_axes(fig, height_px)
    ax_pos = (0.535, 0.12, .17, 0.75) if with_text else (0.03, 0.03, .94, .94)

    return fig, ax_pos
splash_screen(height_px=110, lw_bars=0.7, lw_grid=0.5, lw_border=1,
          rgrid=[1, 3, 5, 7], with_text=True)
plt.show()



root.mainloop()
