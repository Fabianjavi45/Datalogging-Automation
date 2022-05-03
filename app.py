from locale import currency
from multiprocessing.connection import wait
import tkinter as tk
from tkinter import filedialog, Text
from turtle import width
from tkinter import *
from PIL import Image, ImageTk
import Autoxl, Main

root = tk.Tk()
root.title('Efficiency Report Maker')
root.attributes('-fullscreen', True)
originalPath=tk.StringVar()
reportPath=tk.StringVar()
dayPath=tk.StringVar()
var1 = tk.IntVar()
var2 = tk.IntVar()
Checkbutton1 = tk.IntVar()  
orgpathSelected= tk.BooleanVar()
reportpathSelected=tk.BooleanVar()
report_Lock=tk.BooleanVar()
daySelected=tk.BooleanVar()
orgpathSelected.set(False)
reportpathSelected.set(False)
daySelected.set(False)
file_name=[]
Org_Label= Label(root)
Res_Label= Label(root)
fieldLabel= Label(root)
orgLabel= Label(root)
resLabel=Label(root)
dayLabel= Label(root)
class App(tk.Tk):

    def __init__(self):
        super().__init__()

        self.originalPath = tk.StringVar()

def get_original_report():
    global Org_Label, orgpathSelected
    orgReport=filedialog.askopenfilename(initialdir="/Users/fabian/Desktop/PRIFB/EFFIENCY REPORTS/Daily Reports", title="Select ORIGINAL Report File", 
                                        filetypes=(("excel", ".xlsx"), ("all files", ".*")))
    originalPath.set(orgReport)
    if(originalPath.get()!=""):
        Org_Label.destroy()
        orgpathSelected.set(True)
        file_name=orgReport.split("/")
        Org_Label=tk.Label(root,text=file_name[len(file_name)-1],font=("Verdana",14),bg="white",fg="black")
        Org_Label.place(x=5, y=215)
              
def get_Efficiency_report():
    global Res_Label
    resReport=filedialog.askopenfilename(initialdir="/Users/fabian/Desktop/PRIFB/EFFIENCY REPORTS/Daily Reports", title="Select EFFICIENCY Report File", 
                                        filetypes=(("excel", ".xlsx"), ("all files", ".*")))
    reportPath.set(resReport)
    if(reportPath!=""):
        Res_Label.destroy()
        reportpathSelected.set(True)
        file_name=resReport.split("/")
        Res_Label=tk.Label(root,text=file_name[len(file_name)-1],font=("Verdana",14),bg="white",fg="black")
        Res_Label.place(x=5, y= 280)

def get_day_of_Report():
    global daySelected
    if(var1.get()==1):
        dayPath.set("L")
    elif(var1.get()==2):
        dayPath.set("M")
    elif(var1.get()==3):
        dayPath.set("W")
    elif(var1.get()==4):
        dayPath.set("J")
    else:
        dayPath.set("V")
    daySelected.set(True)

def start():
    global fieldLabel, orgLabel, resLabel, dayLabel, orgpathSelected, Res_Label
    if(orgpathSelected.get()!=True or reportpathSelected.get()!=True or daySelected.get()!=True):
        fieldLabel.destroy()
        fieldLabel=tk.Label(root,text="Please select required fields * ",font=("Verdana",14),bg="#F88486",fg="black",relief="solid")
        fieldLabel.place(x=180,y=160)
        if(orgpathSelected.get()!=True):
            orgLabel=tk.Label(root,text=" * ",font=("Verdana",18),bg="white",fg="#F88486")
            orgLabel.place(x=180, y=185)
        else:
            orgLabel.destroy()
        if(reportpathSelected.get()!=True):
            resLabel=tk.Label(root,text=" * ",font=("Verdana",18),bg="white",fg="#F88486")
            resLabel.place(x=195, y=240)
        else:
            resLabel.destroy()
        if(daySelected.get()!=True):   
            dayLabel=tk.Label(root,text=" * ",font=("Verdana",18),bg="white",fg="#F88486") 
            dayLabel.place(x=240, y=310)
        else:
            dayLabel.destroy()
    else:
        Main.main(originalPath.get(),reportPath.get(),dayPath.get())
        #autoxl.main(originalPath.get(),reportPath.get(),dayPath.get())
        orgpathSelected.set(False)
        if(Checkbutton1.get()==0):
            reportpathSelected.set(False)
        daySelected.set(False)
        label=tk.Label(root,text="Effiency Report succesfully made!",font=("Verdana",18),bg="white",fg="black")
        label.pack()
        fieldLabel.destroy()
# Create a photoimage object of the image in the path

image1 = Image.open("res/EfficiencyReportMakerBG.png")
test = ImageTk.PhotoImage(image1)
label1 = tk.Label(image=test)
label1.image = test
label1.place(x = -5, y = -10)

label=tk.Label(root,text=" Efficiency Report Maker (Version 0.1)",font=("Verdana",28),bg="white",fg="black")
label.pack()

#Buttons-----------------------------------------------------------------------------------------------------------------------
#-----------Original Report Search Button-----------#
tk.Button(root, text="Original Report",font=("Verdana",18), fg="black", bg="#F88486", command=get_original_report).place(x=5, y=180)
#-----------Efficiency Report Search Button-----------#
tk.Button(root, text="Efficiency Report", font=("Verdana",18), fg="black", bg="#FFDF13", command=get_Efficiency_report).place(x=5, y=240)
#-----------Day submitting Button-----------#
tk.Button(root, text="Select Day of Report", font=("Verdana",18), fg="black", bg="#FFDF13").place(x=5, y=310)
#Buttons-----------------------------------------------------------------------------------------------------------------------
#-----------INPUT-----------------------#
Radiobutton(root, text='Lunes', variable=var1,bg="#F88486",value=1,command=get_day_of_Report).place(x=5,y=345)
Radiobutton(root, text='Martes', variable=var1,bg="#F88486",value=2,command=get_day_of_Report).place(x=80,y=345)
Radiobutton(root, text='Miercoles', variable=var1,bg="#F88486",value=3,command=get_day_of_Report).place(x=155,y=345)
Radiobutton(root, text='Jueves', variable=var1,bg="#F88486",value=4,command=get_day_of_Report).place(x=245,y=345)
Radiobutton(root, text='Viernes', variable=var1,bg="#F88486",value=5,command=get_day_of_Report).place(x=320,y=345)
#Radiobutton(root, text='Lock Eff. Report', variable=var2,bg="#84C8F8",value=1,command=lock_Report()).place(x=240,y=315)
lockReport=Checkbutton(root, text='Lock Eff. Report',bg="#84C8F8", onvalue = 1, offvalue = 0,variable=Checkbutton1).place(x=230,y=245)
CreateFile = tk.Button(root, text="Make Report", padx=10, pady=5, fg="black", bg="#FFDF13", command=start).place(x=5,y=375)


root.mainloop()