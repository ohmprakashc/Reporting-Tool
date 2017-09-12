# -*- coding: utf-8 -*-
"""
Created on Mon Aug 21 18:35:48 2017
Frontend for the application
@author: Ohmprakash
"""

from tkinter import *
from tkinter import filedialog

window=Tk()

window.title("Reporting tool")

def open_folder():
    window.directory=filedialog.askdirectory() #Gets directory name
    #window.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel files","*.xlsx"),("all files","*.*"))) # Gets file name
    folder_name.insert(END,window.directory) #Writes directory path to folder name box
    #print (window.directory)
    #print ("Success")

#Displays folder path
folder_value=StringVar()
folder_name=Entry(window,textvariable=folder_value)
folder_name.grid(row=0, column=0)

#Button to choose folder path
folder_button=Button(window,text="Choose Folder",command=open_folder)
folder_button.grid(row=0,column=1)

#Button to extract required details in excel files and delete information that is not requried

preprocess_button=Button(window,text="Preprocess data")
preprocess_button.grid(row=1,column=0)


#Button to extract required details in excel files and delete information that is not requried
update_button=Button(window,text="Update DB")
update_button.grid(row=1,column=1)

#Dropdown menu example taken from below web link
#https://pythonspot.com/en/tk-dropdown-example/ for drop down list

#Dropdown list to show BU names
BU_name=StringVar(window)
BU_choice=("All","PEI","Industry","DCE","ITBu","BIPBOP")
BU_name.set("All")
BU_pop_up_menu=OptionMenu(window,BU_name,*BU_choice)
Label(window,text="Choose BU").grid(row=2,column=0)
BU_pop_up_menu.grid(row=2,column=1)

#Dropdown list to show skill names
Skill_name=StringVar(window)
Skill_choice=("All","EM","ESS","TechPub","Simulation","V&V","Software")
Skill_name.set("All")
Skill_pop_up_menu=OptionMenu(window,Skill_name,*Skill_choice)
Label(window,text="Choose Skill").grid(row=3,column=0)
Skill_pop_up_menu.grid(row=3,column=1)

#Dropdown list to show Location names
Location_name=StringVar(window)
Location_choice=("All","SEI","TechM")
Location_name.set("All")
Location_pop_up_menu=OptionMenu(window,Location_name,*Location_choice)
Label(window,text="Choose Location").grid(row=4,column=0)
Location_pop_up_menu.grid(row=4,column=1)

#Dropdown list to show report names
"""
Spend_name=StringVar(window)
Spend_choice=("All","Spend","Headcount","Supplier","Seat cost")
Spend_name.set("All")
Spend_pop_up_menu=OptionMenu(window,Spend_name,*Spend_choice)
Label(window,text="Choose Spend").grid(row=5,column=0)
Spend_pop_up_menu.grid(row=5,column=1)
"""

#Dropdown list to show Month names
"""
Starts at row 5, but displays at row 5 since report name is commented. 
Remove comment if row 5 is enabled
"""
Month_name=StringVar(window)
Month_choice=("All","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
Month_name.set("All")
Month_pop_up_menu=OptionMenu(window,Month_name,*Month_choice)
Label(window,text="Choose Month").grid(row=6,column=0)
Month_pop_up_menu.grid(row=6,column=1)

#Dropdown list to show Year names
Year_name=StringVar(window)
Year_choice=("All",2012,2013,2014,2015,2016,2017,2018,2019,2020)
Year_name.set("All")
Year_pop_up_menu=OptionMenu(window,Year_name,*Year_choice)
Label(window,text="Choose Year").grid(row=7,column=0)
Year_pop_up_menu.grid(row=7,column=1)

#Generates report
generate_button=Button(window,text="Generate Report")
generate_button.grid(row=8,column=1)

window.mainloop()