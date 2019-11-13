'''

'''

#======================
# imports
#======================


from tkinter import ttk, scrolledtext, Menu, Spinbox, \
                    filedialog, messagebox 
import sys                    
from tkinter import  Entry,Label,END,Text
from tkinter import  StringVar,IntVar               
import pathlib
import pandas as pd
import schedule
import threading
import time
from threading import Thread, Lock
_db_lock = Lock()  

from arci_ToolTip_refact import createToolTip       




#from Resources import I18N


class different_tabs():
    def __init__(self,window):
        self.window=window
        
        
    def populate_tab_input(self):
        
        #======================================================================================================
        # allsensors_frame
        #======================================================================================================
        
        allsensors_frame = ttk.LabelFrame(self.tab_input, text='Assigning Sensor names ')
        allsensors_frame.grid(column=0, row=0, padx=8, pady=4)
        
         #Sensor Number
        sensorvar = StringVar(value="Run_1")
        self.sensorentry=Entry(allsensors_frame, textvariable=sensorvar)
        self.sensorentry.grid(row=0,column=1, padx=8, pady=4, sticky='W')
        createToolTip(self.sensorentry, 'Name of the run which can go for 6 days, Enter without spaces.')
        #label for Sensor Number
        ttk.Label(allsensors_frame, text="Run name:").grid(row=0,column=0, padx=8, pady=4, sticky='W')
        
        
        self.entries = []
        for i,sensornumber in enumerate(list(range(101,110))):
            ttk.Label(allsensors_frame, text=str(sensornumber)+":").grid(column=0, row=i+1, padx=8, pady=4, sticky='W')
            intravar=StringVar(value=str(sensornumber))
            single_entry=Entry(allsensors_frame,textvariable=intravar)
            self.entries.append(single_entry)
            single_entry.grid(column=1, row=i+1, padx=8, pady=4, sticky='W')
        
        #======================================================================================================
        # parameters_frame
        #======================================================================================================
        
        
        
        
        parameters_frame = ttk.LabelFrame(self.tab_input, text='Experimental Parameters')
        parameters_frame.grid(column=1, row=0, padx=8, pady=4)
        
        
        ttk.Label(parameters_frame, text="No. of Sensors").grid(column=0, row=0, padx=8, pady=4, sticky='W')
        
        self.combo_num_sensor = ttk.Combobox(parameters_frame, width=12)
        self.combo_num_sensor.config(value = ("1", "2", "3", "4", "5", "6", "7", "8", "9"))
        self.combo_num_sensor.set("6")
        self.combo_num_sensor.grid(column=1, row=0, padx=8, pady=4, sticky='W')
        #numberChosen.current(0)
                
        #Number of Temperature runs for L1, L2 and L3
        temprunsvar=StringVar(value="9")  # temporary variable
        self.temprunsentry=Entry(parameters_frame,width=15,textvariable=temprunsvar)
        self.temprunsentry.grid(row=2,column=1, padx=8, pady=4, sticky='W')
        #Label for number of temperature runs
        ttk.Label(parameters_frame,text="No. of temp runs:").grid(row=2, column=0, padx=8, pady=4, sticky='W')
        
        
        
        #Drop down list for the days
        ttk.Label(parameters_frame, text="Day of run:").grid(column=0, row=4, padx=8, pady=4, sticky='W')
        self.combo = ttk.Combobox(parameters_frame, width=12)
        self.combo.grid(row=4, column=1, padx=8, pady=4, sticky='W')
        self.combo.config(value = ('L1', 'L2', 'L3', 'D4', 'D5','D6','set values'))
        self.combo.set('L1')
        
        #Drop down for internal or external files for Analysis 
        #If internal the location is read from the previously run code
        #else the location is read from the external 
        ttk.Label(parameters_frame, text="Analysis type:").grid(column=0, row=6, padx=8, pady=4, sticky='W')
        self.analysiscombo = ttk.Combobox(parameters_frame, width=12)
        self.analysiscombo.grid(row=6, column=1, padx=8, pady=4, sticky='W')
        self.analysiscombo.config(value = ('Internal', 'External'))
        self.analysiscombo.set('Internal')

        
        #Button for Validation
        buttonschedule = ttk.Button(parameters_frame,width=18,text='Finalize Scheduling', command=self.schedule_all_events)
        buttonschedule.grid(row=0, column=2)
        
        
        #Button for update
        buttonupdate = ttk.Button(parameters_frame,width=18,text='Update',command=self.update)
        buttonupdate.grid(row=2, column=2)
        
        
        
        
        # Button for Acquisition
        self.buttonacquisition =ttk.Button(parameters_frame,text='Acquisition',width=18, command=self.window.acquisition.completeacquisition)
        self.buttonacquisition.grid(row=4, column=2)

        
        #Button to quit
        self.buttonquit=ttk.Button(parameters_frame, text="Quit",width=18,command=self.window.analysis_obj._quit)
        self.buttonquit.grid(row=6, column=2)
        
        
        for child in parameters_frame.winfo_children():
            child.grid_configure(padx=12, pady=4)
       
         #======================================================================================================
        #self.widgetFrame for analysis check boxes
        #======================================================================================================
        
        
        
        
        # We are creating a container frame to hold all other widgets -- Tab2
        self.widgetFrame = ttk.LabelFrame(self.tab_input, text="Analysis with")
        self.widgetFrame.grid(column=2, row=0, padx=8, pady=20)

        self.vars = []
        self.picks=['L1', 'L2', 'L3', 'D4', 'D5','D6']
        self.vars = []
        for num,pick in enumerate(self.picks):
            var = IntVar()
            chk = ttk.Checkbutton(self.widgetFrame,  text=pick, variable=var)
            createToolTip(chk, 'Donot select D5 alone without D4')
            chk.grid(column=num+1, row=0, sticky="W" )
            self.vars.append(var)
                 
         
       # Button for Analysis
        self.buttonanalysis = ttk.Button(self.widgetFrame,text='Analyze',width=18, command=lambda:Thread(target=self.window.analysis_obj.analysis_alldays).start())
        self.buttonanalysis.grid(row=0, column=7)  
        
        
        
        
        
        # Add some space around each label
        for child in self.widgetFrame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        
         #======================================================================================================
        #timercountdown_frame
        #======================================================================================================
        
        information_frame = ttk.LabelFrame(self.tab_input, text="Display Information")
        information_frame.grid(column=0, row=4, padx=8, pady=20)
        
        # Button for opening a file
        self.buttonopenfile = ttk.Button(information_frame,text=self.window.defaultlocation, command=self.openfile, width=50)
        self.buttonopenfile.grid(row=2, column=0)
        self.labelopenfile=ttk.Label(information_frame,text="File location that is read")
        self.labelopenfile.grid(row=1, column=0)
        
        
        
#        #White Space area
#        self.output =Text(information_frame)
#        self.output.grid(row=5,column=2)
        
        style = ttk.Style()
        style.configure("BW.TLabel", foreground="black", background="white")

        l1 = ttk.Label(information_frame, text="Smart Trader Portal", font=("Tahoma", 12, 'bold')).grid(row=3, column=0, padx=8, pady=4, sticky='W')
        self.dtbase_value = StringVar()
        
        
        self.dtbaselbl = ttk.Label(information_frame, textvariable = self.dtbase_value)
        self.dtbaselbl.grid(row=4, column=0, padx=8, pady=4, sticky='W')
        
        self.dtbase_value.set(': connection established')
        self.dtbaselbl.config(foreground="SpringGreen",background="white",font=("Tahoma", 12, 'bold'))
        self.window.master.update_idletasks()
#-------------------------------------------------       
    

    def openfile(self):
        print(messagebox.askquestion(title='Opening file', message='Do you want open file ?'))
        filechosen = filedialog.askopenfile()
        print(filechosen.name)
        self.defaultlocation=filechosen.name
        #defaultlocation
        #self.filename=StringVar(value=filechosen.name)
        print(self.defaultlocation)
        self.buttonopenfile.configure(text=self.defaultlocation)
        #self.entry = TK.Entry(textvariable=self.filename)
        #self.output.insert(END,self.defaultlocation)
    
    #Create a scheduler in paralel mode
    def schedule_all_events(self):
        text_df=pd.read_csv("text.txt",header=None)
        lst_start=text_df.loc[:,0]
        lst_end=text_df.loc[:,1]
        lst_day=text_df.loc[:,2]
        temp=map(self.convert24, lst_start)
        lst_start_24format=list(temp)
        temp=map(self.convert24, lst_end)
        lst_end_24format=list(temp) 
        lstofcommands=[]
        
        #Preparing string for exec function for start time
        for day,start in zip(lst_day,lst_start_24format):
            stringtomake="schedule.every()."+str(day).lower()+".at("+'"'+str(start)+'"'+").do(self.run_threaded, self.window.espec_comm.put_on_espec)"
            print(stringtomake)
            lstofcommands.append(stringtomake)
        
        
        #Preparing string for exec function for stop time
        for day,end_time in zip(lst_day,lst_end_24format):
            stringtomake="schedule.every()."+str(day).lower()+".at("+'"'+str(end_time)+'"'+").do(self.run_threaded, self.window.espec_comm.put_off_espec)"
            print(stringtomake)
            lstofcommands.append(stringtomake)
        
        #schedule.every().saturday.at("02:55").do(self.run_threaded, self.window.espec_comm.put_off_espec)
        
        #schedule.every().friday.at("23:52").do(self.run_threaded, self.job)
        for command in lstofcommands:
            exec(command)
            
        job_thread =Thread(target=self.runningpending, daemon=True)
        job_thread.start()   
        
        
    def job(self):
        print("I'm running on thread %s" % threading.current_thread())
        time.sleep(160)
    
    def runningpending(self):
        
        while True:
            schedule.run_pending()
            time.sleep(1)

    def run_threaded(self,job_func):
        job_thread =Thread(target=job_func, daemon=True)
        job_thread.start()   
        
    def convert24(self,str1): 
      
        # Checking if last two elements of time 
        # is AM and first two elements are 12 
        if str1[-2:] == "AM".lower() and str1[:2] == "12": 
            return "00" + str1[2:-2] 
              
        # remove the AM     
        elif str1[-2:] == "AM".lower(): 
            return str1[:-2] 
          
        # Checking if last two elements of time 
        # is PM and first two elements are 12    
        elif str1[-2:] == "PM".lower() and str1[:2] == "12": 
            return str1[:-2] 
              
        else: 
              
            # add 12 to hours and remove PM 
            return str(int(str1[:2]) + 12) + str1[2:5] 
    
    
    
    
    
     
        
        
    
    def update(self):
        #delay after a click
        self.general_time_delay=200/1000 # in milliseconds
        #delay after a winwait etc.
        self.more_time_delay=500/1000 # in milliseconds





        self.no_channels=int(self.combo_num_sensor.get())
        self.sensor_input_set=self.sensorentry.get()
        print("Sensor run name is  {}".format(self.sensor_input_set))
        
        self.channelnames=[]
        for item in self.entries:
            self.channelnames.append(item.get())
            
        self.channelnames=self.channelnames[0:self.no_channels]
        for channels in self.channelnames:
            print("channel names are:{}".format(channels))
            
        
        
        
        self.tempruns_set=int(self.temprunsentry.get())
        print("Number of temperature runs for L1,L2 and L3 are {}".format(self.tempruns_set))
        self.comboday_set=self.combo.get()
        print("The day you want to run is {}".format(self.comboday_set))
        self.tempruns_D45_set=3
        print("Number of temperature runs for Day4 and Day5 are {}".format(self.tempruns_D45_set))
       
        self.analysiscombo_set=self.analysiscombo.get()
        print("The type of analysis you want is {} ".format(self.analysiscombo_set))


        print("Validating the Path and existence of other files::")
        
        
        self.states_checkbox=list(map((lambda var: var.get()), self.vars))
        for day,state in zip(self.picks,self.states_checkbox):
            print(day,state)
        
        self.ones_days=[day for day,state in zip(self.picks,self.states_checkbox) if state == 1]
        self.zero_days=[day for day,state in zip(self.picks,self.states_checkbox) if state == 0]
        
        if ("D5" in self.ones_days and not "D4" in self.ones_days):
            print ("Please select D4 also if you are selecting D5 \n")
            self.window.master.destroy()
            sys.exit(0)
        
        
        
        
        
        self.root_folder="D:/Auto/"
        #Location where resources can be found
        self.resourcelocation=self.root_folder+'resources/'
        self.datalogger_title=u'Configuration - 2 - BenchLink Data Logger 3'
        
        self.defaultlocation=self.root_folder
        
        
        # All Paths  needs to be here as sensorname is part of the path
        #Base bolder where all the folders shall be made
        
        self.base_folder = "D:/Auto/" + self.sensor_input_set +'/'
        self.base_folder_path= pathlib.Path(self.base_folder)

        #default location where the datalogger saves the file for each scan
        self.dataloggerlocation="C:/Users/PRASAD/Documents/"
        #default path where the datalogger saves the file for each scan
        self.path_to_datalogger_default=pathlib.Path(self.dataloggerlocation)

        

        #Location where the Analysis is done
        self.analysislocation=self.base_folder+'analysis_folder/'

        #raw text or output  raw_text.txt file location
        self.raw_text_path=pathlib.Path(self.base_folder+'raw_text_'+self.comboday_set+'.txt')
        self.combolocation=self.base_folder + self.comboday_set +'/'
    
        print("Updated Successfully")
    
    
    
    
    

        
    def createtabs(self):
        # Tab Control introduced here --------------------------------------
        tabControl = ttk.Notebook(self.window.master)     # Create Tab Control

        
        self.tab1= ttk.Frame(tabControl)
        self.tab2= ttk.Frame(tabControl)
        self.tab3= ttk.Frame(tabControl)
        tabControl.add(self.tab1, text="schedule espec")
        tabControl.add(self.tab2, text="retreive schedule")
        tabControl.add(self.tab3, text="delete schedule")
        
        self.tab_input = ttk.Frame(tabControl)            # Create a tab
        tabControl.add(self.tab_input, text='Input')    # Add the tab -- COMMENTED OUT FOR CH08
        tabControl.pack(expand=1, fill="both")  # Pack to make visible


       
    


#if __name__ == '__main__':
#    #======================
#    # Start GUI
#    #======================
#    window = Window()
##    print(oop.log)
###     oop.log.setLoggingLevel(oop.level.DEBUG)
##    oop.log.setLoggingLevel(oop.level.MINIMUM)
##    oop.log.writeToLog('Test message')
#    window.master.mainloop()
