# -*- coding: utf-8 -*-
"""
Created on Tue Oct 15 00:46:01 2019

@author: sairamtvv
"""

"""imports"""
#++++++++++++++++++++++++++++++++++++++++++++
from datetime import datetime
from tkinter import messagebox as tkMessageBox
from tkinter import font as tkFont
from tkinter import StringVar,ttk
#from tkinter import *
from tkinter import Menu,Text,Button,INSERT,END,X,PhotoImage
import os
from prettytable import PrettyTable
from PIL import ImageTk, Image
import pandas as pd


import threading
import time
import schedule

#++++++++++++++++++++++++++++++++++++++++++++

#++++++++++++++++++++++++++++++++++++++++++++
"""creating the window, menubar and title"""
class Espec_Scheduler():
    def __init__(self,window):
        #creating instance
        self.window=window
        self.root = window.master
        self.tkMessageBox=tkMessageBox
        self.ttk=ttk
        #using prettytable to create the tables for start and end time and event
        headerFormat = ["Start Time","EndTime","Day","Comment"]
        #self.mains=PrettyTable(headerFormat)
        self.main1=PrettyTable(headerFormat)
        self.main2=PrettyTable(headerFormat)
        self.main3=PrettyTable(headerFormat)
        self.main4=PrettyTable(headerFormat)
        self.main5=PrettyTable(headerFormat)
        self.main6=PrettyTable(headerFormat)
        self.main7=PrettyTable(headerFormat)
        self.customFont = tkFont.Font(family="Helvetica", size=12)       

#        #initializing functions
#        self.root.resizable(0,0)
#        self.createMenu()
#        self.MainTabs()
#        self.TabContent()
#        self.retrieveTab()
#        self.deleteEvent()
#++++++++++++++++++++++++++++++++++++++++++

#==================================================================        
    
    
    
    
    
    
    
    #creating the menu bar
    def createMenu(self):
        self.menuBar=Menu(self.root)
        self.root.configure(menu= self.menuBar)
            ##
        self.FileMenu = Menu(self.menuBar,tearoff=0)
        self.FileMenu.add_command(label="New")
        self.FileMenu.add_command(label="Exit", command=self.root.destroy)
        self.HelpMenu= Menu(self.menuBar,tearoff=0)
        self.HelpMenu.add_command(label="About", command=self._msgBox)
        self.menuBar.add_cascade(label="File", menu=self.FileMenu)
        self.menuBar.add_cascade(label="Help", menu=self.HelpMenu)
        
    def _msgBox(self):
        self.tkMessageBox.showinfo('About the App','This App is designed to create a schedule for the days of the week and the time with an hour between between the times.The schedule can be accessed by the days of the week')
#=====================================================================

#    #creating the tabs
    def MainTabs(self):
        
        """Adding tabs to make navigation easy"""
        self.AddingEventTab=ttk.Notebook(self.root)
        self.tab1= ttk.Frame(self.AddingEventTab)
        self.tab2= ttk.Frame(self.AddingEventTab)
        self.tab3= ttk.Frame(self.AddingEventTab)
        self.AddingEventTab.add(self.tab1, text="ADD EVENT")
        self.AddingEventTab.add(self.tab2, text="RETRIEVE EVENT")
        self.AddingEventTab.add(self.tab3, text="DELETE EVENT")
        self.AddingEventTab.pack(expand=1,pady=5, fill="both")

#**********************************************************************************************************
    """Adding content to the 'Add event tab'"""
    def populate_tab_schedule_espec(self):
        #first choosing day
        root1= ttk.LabelFrame(self.window.diff_tabs.tab1, text="Add Event")
        root1.grid(column=0, row=0,padx=8,pady=4,sticky="W")
        Day=ttk.Label(root1, text="Please Choose The Day").grid(column=0,row=0,sticky='W')
        self.days=StringVar()
        self.times=StringVar()
        self.times2=StringVar()
        self.bgenda=StringVar()
        option1=self.ttk.Combobox(root1, width=20, textvariable=self.days,state="readonly")
        option1['values']=("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
        option1.grid(column=0,row=1,sticky='W')
        option1.current(0)
#**************************************************************************************************************

        ###choosing start time
        Time=ttk.Label(root1, text="Please Choose a Time").grid(column=0,row=2,sticky='W')
        startTime=ttk.Label(root1, text="Start Time").grid(column=0,row=3,sticky='W')
        EndTime=ttk.Label(root1, text="End Time").grid(column=2,row=3)

        ###Drop down days of the week
        Timedrop=ttk.Combobox(root1, width=20, textvariable=self.times,state="readonly")
        Timedrop['values']=("12:00am","01:00am","02:00am","03:00am","04:00am","05:00am","06:00am","07:00am","08:00am","09:00am",
                            "10:00am","11:00am","12:00pm","01:00pm","02:00pm","03:00pm","04:00pm","05:00pm","06:00pm","07:00pm","08:00pm","09:00pm","10:00pm","11:00pm")
        Timedrop.grid(column=0,row=4,sticky='W')
        Timedrop.current(0)
        
        TimeEndDrop=ttk.Combobox(root1,width=20, textvariable=self.times2,state="readonly")
        
        TimeEndDrop['values']=("12:00am","01:00am","02:00am","03:00am","04:00am","05:00am","06:00am","07:00am","08:00am","09:00am",
                            "10:00am","11:00am","12:00pm","01:00pm","02:00pm","03:00pm","04:00pm","05:00pm","06:00pm","07:00pm","08:00pm","09:00pm","10:00pm","11:00pm")
        TimeEndDrop.grid(column=2,row=4)
        TimeEndDrop.current(0)
        ttk.Label(root1,text="Please enter a short comment if you like").grid(column=0,row=7,pady=10,sticky='W')
        #ttk.Label(root1,text="Examples are class, meeting, work, workout").grid(column=0,row=7,sticky='W')

        self.Agenda=ttk.Entry(root1,width=50,textvariable=self.bgenda)
        self.Agenda.grid(column=0,row=8,pady=5,sticky='W')
        showbiz2=Text(root1,width=85,height=15)
        showbiz2.grid(column=0,row=9)

        def viewPrint():
            showbiz2.delete(1.0,END)
            tkMessageBox.showinfo("View","The View button shows all the Events")
            if os.stat("text.txt").st_size==0:
                showbiz2.insert(INSERT,"No Event Found")
            else:
                with open ("text.txt","r")as file_new:
                    contents=file_new.read()
                    showbiz2.insert(INSERT,contents)
                    showbiz2.see(INSERT)
        def clean():
            showbiz2.delete(1.0,END)
            tkMessageBox.showinfo("Clear","Please note that the clear button only cleans the display field but not the schedule. Please use 'DELETE EVENT' to delete any event")
            
                
        creat=Button(self.window.diff_tabs.tab1,text="CREATE  ",bg='#4d79ff',fg='white',font=("Helvetica",14),command=self.addElement)
        creat.grid(column=0,ipadx=62,pady=5,row=10,sticky="W")
        view=Button(self.window.diff_tabs.tab1, text="All Event",bg='#ff5050',fg='white',font=("Helvetica",14),command=viewPrint)
        view.grid(column=0,ipadx=70,row=11,pady=5,sticky="W")
        clear=Button(self.window.diff_tabs.tab1,text="CLEAR   ",bg='#ffc34d',fg='white',font=("Helvetica",14),command=clean)
        clear.grid(column=0,ipadx=66,row=12,pady=5,sticky="W")
    def addElement(self):
        def allEvent():
            tkMessageBox.showinfo("Created","Event Created")
            self.Agenda.delete(0,'end')
            #self.mains.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("text.txt","a")
            #file_new.write(str("Created: " + datetime.now().strftime(' %Y-%m-%d %H:%M:%S'))+'\n')
            file_new.write(f"{self.times.get()},{self.times2.get()},{self.days.get()}")
            file_new.write("\n")
#            if self.days.get()=="Monday":
#                file_new.write(str(self.main1)+'\n')
#            elif self.days.get()=="Tuesday":
#                file_new.write(str(self.main2)+'\n')
#            elif self.days.get()=="Wednesday":
#                file_new.write(str(self.main3)+'\n')
#            elif self.days.get()=="Thursday":
#                file_new.write(str(self.main4)+'\n')
#            elif self.days.get()=="Friday":
#                file_new.write(str(self.main5)+'\n')
#            elif self.days.get()=="Saturday":
#                file_new.write(str(self.main6)+'\n')
#                
#            elif self.days.get()=="Sunday":
#                file_new.write(str(self.main7))
            file_new.close()
        def monday():
            self.main1.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("monday.txt","w")
            file_new.write(str(self.main1))
            file_new.close()
            allEvent()
        def tuesday():
            self.main2.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("tuesday.txt","w")
            file_new.write(str(self.main2))
            file_new.close()
            allEvent()

        def wednesday():
            self.main3.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("wednesday.txt","w")
            file_new.write(str(self.main3))
            file_new.close()
            allEvent()

        def thursday():
            self.main4.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("thursday.txt","w")
            file_new.write(str(self.main4))
            file_new.close()
            allEvent()

        def friday():
            self.main5.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("friday.txt","w")
            file_new.write(str(self.main5))
            file_new.close()
            allEvent()

        def saturday():
            self.main6.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("saturday.txt","w")
            file_new.write(str(self.main6))
            file_new.close()
            allEvent()

        def sunday():
            self.main7.add_row([self.times.get(),self.times2.get(),self.days.get(),self.bgenda.get()])
            file_new=open("sunday.txt","w")
            file_new.write(str(self.main7))
            file_new.close()
            allEvent()

        if self.days.get()=="Monday":
            monday()
        elif self.days.get()=="Tuesday":
            tuesday()
        elif self.days.get()=="Wednesday":
            wednesday()
        elif self.days.get()=="Thursday":
            thursday()
        elif self.days.get()=="Friday":
            friday()
        elif self.days.get()=="Saturday":
            saturday()
        elif self.days.get()=="Sunday":
            sunday()
            

#************************************************************************************************************************
        """Retrieving the Envents"""
    def populate_tab_retrieveschedule(self):
        root2=ttk.LabelFrame(self.window.diff_tabs.tab2, text="Retrieve")
        greetings=ttk.Label(self.window.diff_tabs.tab2,text="Hello there!Please click on the day you want to retrieve",font=self.customFont)
        greetings.grid(column=0,row=0,pady=5)
        
        def printAllEvent():
            showbiz.delete(1.0,END)
            file_new=open("text.txt",'a')
            if os.stat("text.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, No Event Found")
            else:
                with open ("text.txt","a")as file_new:
                    file_new=open("text.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printMonday():
            showbiz.delete(1.0,END)
            file_new=open("monday.txt","a")
            if os.stat("monday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Monday is Empty")
            else:
                with open ("monday.txt","a")as file_new:
                    file_new=open("monday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printTuesday():
            showbiz.delete(1.0,END)
            file_new=open("tuesday.txt","a")
            if os.stat("tuesday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Tuesday is Empty")
            else:
                with open ("tuesday.txt","a")as file_new:
                    file_new=open("tuesday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printWednesday():
            showbiz.delete(1.0,END)
            file_new=open("wednesday.txt",'a')
            if os.stat("wednesday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Wednesday is Empty")
            else:
                with open ("wednesday.txt","a")as file_new:
                    file_new=open("wednesday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printThursday():
            showbiz.delete(1.0,END)
            file_new=open("thursday.txt",'a')
            if os.stat("thursday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Thursday is Empty")
            else:
                with open ("thursday.txt","a")as file_new:
                    file_new=open("thursday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printFriday():
            showbiz.delete(1.0,END)
            file_new=open("friday.txt",'a')
            if os.stat("friday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Friday is Empty")
            else:
                with open ("friday.txt","a")as file_new:
                    file_new=open("friday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printSaturday():
            showbiz.delete(1.0,END)
            file_new=open("saturday.txt",'a')
            if os.stat("saturday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Saturday is Empty")
            else:
                with open ("saturday.txt","a")as file_new:
                    file_new=open("saturday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
        def printSunday():
            showbiz.delete(1.0,END)
            file_new=open("sunday.txt",'a')
            if os.stat("sunday.txt").st_size == 0:
                showbiz.insert(INSERT,"Sorry, Sunday is Empty")
            else:
                with open ("sunday.txt","a")as file_new:
                    file_new=open("sunday.txt")
                    contents=file_new.read()
                    showbiz.insert(INSERT,contents)
                    showbiz.see(INSERT)
       
        monday=Button(self.window.diff_tabs.tab2, text="Monday",bg="#4d79ff",fg="white",font=("Helvetica",14),command=printMonday)
        monday.grid(column=0,row=2,ipadx=18,pady=5)
        tuesday=Button(self.window.diff_tabs.tab2, text="Tuesday",bg="#ff5050",fg="white",font=("Helvetica",14),command=printTuesday)
        tuesday.grid(column=1,row=2,ipadx=11,pady=5)
        wednesday=Button(self.window.diff_tabs.tab2, text="Wednesday",bg="#ffc34d",fg="white",font=("Helvetica",14),command=printWednesday)
        wednesday.grid(column=0,row=3,pady=5)
        thursday=Button(self.window.diff_tabs.tab2, text="Thursday",bg="#4d79ff",fg="white",font=("Helvetica",14),command=printThursday)
        thursday.grid(column=1,row=3,ipadx=10,pady=5)
        friday=Button(self.window.diff_tabs.tab2, text="Friday",bg="#009900",fg="white",font=("Helvetica",14),command=printFriday)
        friday.grid(column=0,row=4,pady=5,ipadx=24)
        saturday=Button(self.window.diff_tabs.tab2, text="Saturday",bg="#ff5050",fg="white",font=("Helvetica",14),command=printSaturday)
        saturday.grid(column=1,row=4,ipadx=15,pady=5)
        sunday=Button(self.window.diff_tabs.tab2, text="Sunday",bg="#ffc34d",fg="white",font=("Helvetica",14),command=printSunday)
        sunday.grid(column=0,row=5,pady=5,ipadx=20)
        showbiz=Text(self.window.diff_tabs.tab2,width=97,height=22)
        showbiz.grid(columnspan=3)
        allEvent=Button(self.window.diff_tabs.tab2, text="All Events",bg="orange",fg="white",font=("Helvetica",14),command=printAllEvent)
        allEvent.grid(column=1,row=5,ipadx=8,pady=5)
#***********************************************************************************************************************
    def populate_tab_deleteschedule(self):
        def deleteMonday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete Monday Event ?")
            if result is True:
                file_new=open("monday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Monday Event(s) Deleted")
        def deleteTuesday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete Tuesday Event ?")
            if result is True:
                file_new=open("tuesday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Tuesday Event(s) Deleted")
            
        def deleteWednesday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete Wednesday Event ?")
            if result is True:
                file_new=open("wednesday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Wednesday Event(s) Deleted")
        def deleteThursday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete thursday Event ?")
            if result is True:
                file_new=open("thursday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Thursday Event(s) Deleted")
        def deleteFriday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete Friday Event ?")
            if result is True:
                file_new=open("friday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Friday Event(s) Deleted")
        def deleteSaturday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete Saturday Event ?")
            if result is True:
                file_new=open("saturday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Saturday Event(s) Deleted")
        def deleteSunday():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete Sunday Event ?")
            if result is True:
                file_new=open("sunday.txt","w").close()
                tkMessageBox.showinfo("Deleted","Sunday Event(s) Deleted")
        def deleteAllEvent():
            result=tkMessageBox.askyesno("delete","Are you sure you want to delete All Events ?")
            files = ["text.txt", "monday.txt", "tuesday.txt", "wednesday.txt","thursday.txt","friday.txt" ,"saturday.txt" ,"sunday.txt"]
            if result is True:
                for file_name in files:
                    file_new=open(file_name,"w").close()
                tkMessageBox.showinfo("Deleted","All Events Deleted")

        monday=Button(self.window.diff_tabs.tab3, text="Delete Monday",bg='#4d79ff',fg='white',font=("Helvetica",14),command=deleteMonday)
        monday.pack(fill=X,padx=100,ipady=10,pady=10)
        tuesday=Button(self.window.diff_tabs.tab3, text="Delete Tuesday",bg='#ff5050',fg='white',font=("Helvetica",14),command=deleteTuesday)
        tuesday.pack(fill=X,padx=100,ipady=10,pady=10)
        wednesday=Button(self.window.diff_tabs.tab3, text="Delete Wednesday",bg='#ffc34d',fg='white',font=("Helvetica",14),command=deleteWednesday)
        wednesday.pack(fill=X,padx=100,ipady=10,pady=10)
        thursday=Button(self.window.diff_tabs.tab3, text="Delete Thursday",bg='#4d79ff',fg='white',font=("Helvetica",14),command=deleteThursday)
        thursday.pack(fill=X,padx=100,ipady=10,pady=10)
        friday=Button(self.window.diff_tabs.tab3, text="Delete Friday",bg='#009900',fg='white',font=("Helvetica",14),command=deleteFriday)
        friday.pack(fill=X,padx=100,ipady=10,pady=10)
        saturday=Button(self.window.diff_tabs.tab3, text="Delete Saturday",bg='#ff5050',fg='white',font=("Helvetica",14),command=deleteSaturday)
        saturday.pack(fill=X,padx=100,ipady=10,pady=10)
        sunday=Button(self.window.diff_tabs.tab3, text="Delete Sunday",bg='#ffc34d',fg='white',font=("Helvetica",14),command=deleteSunday)
        sunday.pack(fill=X,padx=100,ipady=10,pady=10)
        allEvent=Button(self.window.diff_tabs.tab3, text="Delete all Events",bg='orange',fg='white',font=("Helvetica",14),command=deleteAllEvent)
        allEvent.pack(fill=X,padx=100,ipady=10,pady=10)

#if __name__ == '__main__':
#    #======================
#    # Start GUI
#    #======================
#    app= EventApp()
#    app.root.mainloop()
#        
