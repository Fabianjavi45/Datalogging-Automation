import tkinter as tk
from tkinter import filedialog, Text
from turtle import width
from tkinter import *
from PIL import Image, ImageTk
import os
import autoxl

root = tk.Tk()
root.title('Efficiency Report Maker')
root.attributes('-fullscreen', True)
file_names=[]

def get_original_report():
    
    orgReport=filedialog.askopenfilename(initialdir="/Users/fabian/Desktop/PRIFB/EFFIENCY REPORTS/Daily Reports", title="Select Report File", 
                                        filetypes=(("excel", ".xlsx"), ("all files", ".*")))
    file_names.append(orgReport)
    print(orgReport)
    file_name=orgReport.split("/")
    label=tk.Label(root,text="Original Report: "+file_name[len(file_name)-1],font=("Verdana",18),bg="white",fg="black")
    label.pack()

def get_Efficiency_report():
    
    resReport=filedialog.askopenfilename(initialdir="/Users/fabian/Desktop/PRIFB/EFFIENCY REPORTS/Daily Reports", title="Select Report File", 
                                        filetypes=(("excel", ".xlsx"), ("all files", ".*")))
    file_names.append(resReport)
    print(resReport)
    file_name=resReport.split("/")
    label=tk.Label(root,text="Efficiency Report: "+file_name[len(file_name)-1],font=("Verdana",18),bg="white",fg="black")
    label.pack()

def get_day_of_Report():
   Day = entry.get()
   file_names.append(Day)
   if(Day=="L"):
       label=tk.Label(root,text="Day of Report: Lunes",font=("Verdana",18),bg="white",fg="black")
   elif(Day=="M"):
       label=tk.Label(root,text="Day of Report: Martes",font=("Verdana",18),bg="white",fg="black")
   elif(Day=="W"):
       label=tk.Label(root,text="Day of Report: Miercoles",font=("Verdana",18),bg="white",fg="black")
   elif(Day=="J"):
       label=tk.Label(root,text="Day of Report: Jueves",font=("Verdana",18),bg="white",fg="black")   
   elif(Day=="V"):
       label=tk.Label(root,text="Day of Report: Viernes",font=("Verdana",18),bg="white",fg="black")     
   label.pack()

def start():
    autoxl.main(file_names[0],file_names[1],file_names[2])
    print("This")
    print("That")
    
# Create a photoimage object of the image in the path
image1 = Image.open("res/EfficiencyReportMakerBG.png")
test = ImageTk.PhotoImage(image1)
label1 = tk.Label(image=test)
label1.image = test
label1.place(x = 0, y = -1)

label=tk.Label(root,text=" Efficiency Report Maker (Version 0.1)",font=("Verdana",28),bg="white",fg="black")
label.pack()

# Position image
#canvas = tk.Canvas(root,height=700, width=700,bg="#FFE49C")
#canvas.pack()

#frame = tk.Frame(root, bg="white")
#frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

CreateFile = tk.Button(root, text="Original Report", padx=10, pady=5, fg="black", bg="#FFDF13", command=get_original_report)
CreateFile.pack()

CreateFile = tk.Button(root, text="Efficiency Report", padx=10, pady=5, fg="black", bg="#FFDF13", command=get_Efficiency_report)
CreateFile.pack()

label=tk.Label(root,text="Day of report",font=("Verdana",14),bg="white",fg="black")
label.pack()

label=tk.Label(root,text="L (Lunes), M (Martes), W (Miercoles), J (Jueves) & V (Viernes)",font=("Verdana",14),bg="white",fg="black")
label.pack()

entry= Entry(root)
entry.config(bg="black")
entry.pack()

submit = Button(root, text="Submit Day of Report", command=get_day_of_Report)
submit.pack()

CreateFile = tk.Button(root, text="Make Report", padx=10, pady=5, fg="black", bg="#FFDF13", command=start)
CreateFile.pack()

root.mainloop()