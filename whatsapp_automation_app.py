try:
    import pywhatkit as pt
    internet=True
except:internet=False
import pyautogui as gu
import time
import webbrowser as wb
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import ttk,Text
import os


win=Tk()
win.geometry("600x500")
win.minsize(600,500)

def browse_file():
    number_comb.delete(0, END)
    name_comb1.delete(0, END)
    file_path = filedialog.askopenfilename(initialdir = "/", title = "Select file", filetypes = (("Text files", ".xlsx"), ("all files", ".xlsx")))
    path_entry.delete(0, END)
    path_entry.insert(0, file_path)

def get_values():
    global df
    c=0
    
    if os.path.exists(path_entry.get())==False:
        warning_label.config(text="check the file path")
        return
    df=pd.read_excel(path_entry.get())

    if name_comb1.get().strip()=="":
         warning_label.config(text="plz select name column")
         return
    if number_comb.get().strip()=="":
         warning_label.config(text="plz select phone number column")
         return
        
    if name_comb1.get() not in list(df.columns):
        warning_label.config(text=f"{name_comb1.get()} column is not in excel sheet")
        return
    if number_comb.get() not in list(df.columns):
        warning_label.config(text=f"{number_comb.get()} column is not in excel sheet")
        return
    
    for i in list(df[name_comb1.get()]):
        if type(i)!=str:
            warning_label.config(text="name coloumn contain numneric values")
            return
    if len(list(df[number_comb.get()])) ==0 or len(list(df[name_comb1.get()]))==0:
        warning_label.config(text="empty coloumn")
        return
        
    if (sum([1 for i in map(str,list(df[number_comb.get()])) if (i.isnumeric() and len(i)==10)])/len(list(df[number_comb.get()])))*100<40:
        warning_label.config(text="phone number column is not correct check excel sheet")
        return
     
    if text_entry.get("1.0","end-1c")=="" or text_entry.get("1.0","end-1c")=="type your message here....." :
        warning_label.config(text="plz enter the message")
        return
    
    
    if internet:
            print("You are connected to the Internet.")
            wb.open("https://web.whatsapp.com/")
            
            while True:
                    if gu.locateCenterOnScreen(r'C:\Users\hp\OneDrive\Desktop\what.png')!=None:
                               time.sleep(1)
                               gu.hotkey("ctrl","w")
                               time.sleep(1)
                               gu.press("enter")
                               
                               break
            for i in range(len(list(df[name_comb1.get()]))):
              if df.loc[i,"status"]!="done": 
                
                  try:
                      phone=str(df.loc[i,number_comb.get()])
                      if len(phone)==10 and phone.isnumeric(): 
                           pt.sendwhatmsg_instantly("+91"+str(df.loc[i,number_comb.get()]),"Hey "+str(df.loc[i,str(name_comb1.get())])+", "+str(text_entry.get("1.0","end-1c")))
                           
                         
                           time.sleep(2)
                           gu.hotkey("ctrl","w")
                           time.sleep(1)
                           gu.press("enter")
                           
                           time.sleep(2)
                           df.loc[i,"status"]="done"
                      else:
                           df.loc[i,"status"]="Not done"
                  except:
                      
                      df.loc[i,"status"]="Not done"
                      
            df.to_excel(path_entry.get(),index=False)
        
    else:
        print("You are not connected to the Internet.")
    
    
    

def clear(event):
    
    if text_entry.get("1.0","end-1c")=="type your message here.....":
        text_entry.delete("1.0","end")
def clear1(event):
    
     number_comb.delete(0, END)
     name_comb1.delete(0, END)
def path_input(event):
    
    global df
    if os.path.exists(path_entry.get()) and str(path_entry.get()).split(".")[-1]=="xlsx":
        print("Excel File Selected")
        df=pd.read_excel(str(path_entry.get()))
        name_comb1.configure(values=tuple(df.columns))
        number_comb.configure(values=tuple(df.columns))
    else:
      print("Selsect excel file")
      name_comb1.configure(values=())
      number_comb.configure(values=())

filename6=PhotoImage(file=r"C:\Users\hp\OneDrive\Desktop\New folder\new_what.png")
background_lable16=Label(win,image=filename6)
background_lable16.place(x=0,y=0)

name_comb1=ttk.Combobox(win,values=[],width=20)
name_comb1.place(relx=0.3,rely=0.41,anchor="center")
name_comb1.bind("<Button-1>",path_input)

warning_label= Label(win,font=('helvetica',11,),text="",fg="#ff0000",borderwidth=0,border=0)
warning_label.place(relx=0.5,rely=0.769,anchor="center")

number_comb=ttk.Combobox(win,values=[],width=20)
number_comb.place(relx=0.73,rely=0.41,anchor="center")
number_comb.bind("<Button-1>",path_input)

path_entry = Entry(win,width =36)
path_entry.place(relx=0.48,rely=0.22,anchor="center")
path_entry.bind("<FocusIn>",clear1)

path_label=Label(win,font=('Times New Roman',12,"bold"),text="File path: ",borderwidth=0)
path_label.place(relx=0.23,rely=0.22,anchor="center")

name_label=Label(win,font=('Times New Roman',12,"bold"),text="Name Coloumn",borderwidth=0)
name_label.place(relx=0.3,rely=0.33,anchor="center")

number_label=Label(win,font=('Times New Roman',12,"bold"),text="Phone Number Coloumn",borderwidth=0)
number_label.place(relx=0.72,rely=0.33,anchor="center")

text_entry = Text(win,width=50,height=8)
text_entry.insert(END,"type your message here.....")

text_entry.place(relx=0.52,rely=0.6,anchor="center")
text_entry.bind("<FocusIn>",clear)



send_b1=Button(win,text="Add",font=('Times New Roman',12,"bold"),width=12,command=get_values)
send_b1.place(relx=0.5,rely=0.85,anchor="center")

browse_b1=Button(win,text="browse",font=('Times New Roman',10,"bold"),width=12,command= browse_file)
browse_b1.place(relx=0.76,rely=0.22,anchor="center")


Title_label=Label(win,font=('Times New Roman',19,"bold"),text="Whatsapp Message Automation",borderwidth=0)
Title_label.place(relx=0.52,rely=0.05,anchor="center")               

win.mainloop()












