from tkinter import *
import tkinter as tk
from tkinter.font import Font
import webbrowser
from tkinter import ttk
from tkinter import filedialog,messagebox
import pyperclip
import csv
import os
import sys
import os.path
import tkinter.scrolledtext as scrolledtext
import sqlite3
import pandas as pd 
from sqlalchemy import create_engine
from openpyxl import Workbook
from pandastable import Table
import openpyxl
import traceback

window =tk.Tk()
main_menu=tk.Menu(window)

#-----------------main gui title-----------------
window.title("PyGems Excel To DataTable")

wi_gui=900
hi_gui=600

wi_scr=window.winfo_screenwidth()
hi_scr=window.winfo_screenheight()

x=(wi_scr/2)-(wi_gui/2)
y=(hi_scr/2)-(hi_gui/2)

window.geometry('%dx%d+%d+%d'%(wi_gui,hi_gui,x,y))



# ALL FRAMES

inputframe=Frame(window)
footer=Frame(window)
inputQuery=Frame(window)
outputdata =Frame(window)
exportbtn = Frame(window)
err=Frame(window)


# label frame : 

fileinput=Frame(window)
qeditor=Frame(window)
tableview = Frame(window)
errorview = Frame(window)





#all function 

def opnlink(url):
    webbrowser.open_new(url)
def daction():
	entry_d.delete(0, 'end')
	
	daction.file_selected = filedialog.askopenfilename(initialdir="/",title='Please Select The file',filetypes=(("Excel file","*.xlsx"),("all files","*.")))
		#os.chdir(chd)
	wb_obj1 = openpyxl.load_workbook(daction.file_selected)
	sheet_obj1 = wb_obj1.active.title 
	l.config(text='Your Active SheetName is :'+str(sheet_obj1))


	try:
		if not daction.file_selected:
			daction.file_selected=entry_d_var.get()
			# print(daction.file_selected)
		else:
			df = pd.read_excel(daction.file_selected)
			# print(df)
			entry_d.insert(0,daction.file_selected)
	except:
		messagebox.showerror("Error", "File is Read Only Or Wrong Directory !\n  Visit : pygems.com For Tutorial")

be=" Note : Active SheetName considered as Your Table Name "
l = Label (err, text =be,bd=0,
	fg='#F4511E',
	font='Times 12',
	width=90, )
l.pack()

def retrieve_input():
	try:
	    retrieve_input.df5=""
	    pd.options.display.float_format = '{:.2f}'.format
	    file=daction.file_selected
	    inputValue=textBox.get("1.0","end-1c")
	    # print(inputValue)
	    wb_obj = openpyxl.load_workbook(file)
	    sheet_obj = wb_obj.active.title 
	    engine = create_engine('sqlite://', echo=False)
	    df = pd.read_excel(file,sheet_name=str(sheet_obj))
	    # df = df.astype(str)
	    df.to_sql(sheet_obj,engine,if_exists='replace',index=False)
	    retrieve_input.df5 = pd.read_sql_query(inputValue,engine)
	    outputdata.pack()
	    pt = Table(outputdata,rows=3,dataframe=retrieve_input.df5,showtoolbar=False, showstatusbar=False)
	    pt.show()

	    err.pack_forget()
	except Exception as e:
		outputdata.pack_forget()
		err.pack()
		l.config(text=str(e))

def expexcel():
	try:
		if retrieve_input.df5.empty:
			messagebox.showerror("Error", "Have Nothing to Export")
		else:
		    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
		    retrieve_input.df5.to_excel (export_file_path, index = False, header=True)
	except:
		messagebox.showerror("Error", "Have Nothing to Export")

def expcsv():
	try:
		if retrieve_input.df5.empty:
			messagebox.showerror("Error", "Have Nothing to Export")
		else:
		    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
		    retrieve_input.df5.to_csv (export_file_path, index = False, header=True)
	except:
		messagebox.showerror("Error", "Have Nothing to Export")


def copy():
	# try:
	a=retrieve_input.df5.to_string(header = True, index = False)
	pyperclip.copy(a)
	# 	# btn_txt.set("Copied")
	# except:
	# 	messagebox.showerror("Error", "Have Nothing To Copy")



lbl = Label(fileinput, text="Select Your Excel File : ",
	bd=0,
	anchor="w",
	fg='#132238',
	font='Times 12',
	width=80,
	)

lbl.pack()








entry_d_var = StringVar()
entry_d=Entry(inputframe,width=100,textvariable=entry_d_var,bg='#dedede')
entry_d_txt = entry_d_var.get()

#------------directory Button----------
button_d=tk.Button(inputframe,relief=RAISED,font=('Times 10 bold'),text='Select File' ,fg='#fcf9ec',bg='#132238',command=daction)

entry_d.pack(side=LEFT,ipady=2)
entry_d.focus()
button_d.pack(side=LEFT,padx=10,ipady=2,pady=13)

lbl1 = Label(qeditor, text="Query Editor ",
	bd=0,
	anchor="w",
	fg='#132238',
	font='Times 12',
	width=80,
	)

lbl1.pack(pady=5)

textBox=scrolledtext.ScrolledText(inputQuery, height=5, width=85,undo=True)
textBox.pack(pady=5)

buttonCommit=Button(inputQuery, text="Execute Query", 
                    command=lambda: retrieve_input(),width=20,relief=RAISED,font=('Times 10 bold'),fg='#fcf9ec',bg='#132238')
#command=lambda: retrieve_input() >>> just means do this when i press the button
buttonCommit.pack(pady=5)

#--------------- copy frame -----
cccbtn=StringVar()
ccclbl = Label(exportbtn, text=" ")
ccclbl.pack(side=LEFT,padx=0)
ccclbtn = Button(exportbtn, text="Export XLSX",textvariable=cccbtn, command=expexcel,width=20,relief=RAISED,font=('Times 10 bold'),fg='#fcf9ec',bg='#132238')
cccbtn.set("Export XLSX")
ccclbtn.pack(side=LEFT)

sccbtn=StringVar()
scclbl = Label(exportbtn, text="    ")
scclbl.pack(side=LEFT,padx=10)
scclbtn = Button(exportbtn, text="Export CSV",textvariable=sccbtn, command=expcsv,width=20,relief=RAISED,font=('Times 10 bold'),fg='#fcf9ec',bg='#132238')
sccbtn.set("Export CSV")
scclbtn.pack(side=LEFT)


lcbtn=StringVar()
cclbl = Label(exportbtn, text="       ")
cclbl.pack(side=LEFT,padx=10)
cclbtn = Button(exportbtn, text="Copy Data", command=copy,textvariable=lcbtn,width=20,relief=RAISED,font=('Times 10 bold'),fg='#fcf9ec',bg='#132238')
lcbtn.set("Copy Data")
cclbtn.pack(side=LEFT)





#----------- status frame ---------
# statusbar =Label(window, text="Click here to visit : pygems.com ",
#  bd=1,
#   relief=SUNKEN,
#    bg="#37474F",
#    fg='#fcf9ec',
#    height=2,
#    font="Times 13",
#    cursor="hand2",
#    width=80
#    )

# statusbar.bind("<Button-1>", lambda e: opnlink("https://pygems.com/"))
# statusbar.pack(side=BOTTOM)



fileinput.pack(pady=5)
tableview.pack(padx=0)
errorview.pack(padx=0)


inputframe.pack(padx=0)
qeditor.pack(padx=0)

inputQuery.pack(padx=0)

exportbtn.pack(pady=20)
err.pack(padx=0)
outputdata.pack()


# footer.pack(pady=0)

#-------------- Frames End --------------------
inputframe.config(bg="#ffffff")

window.mainloop()