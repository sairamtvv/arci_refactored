'''
Created on July 26, 2016
Section 8
@author: test
'''
#======================
# imports
#======================
from ttkthemes import themed_tk  as tk
from tkinter import ttk, scrolledtext, Menu, Spinbox, \
                    filedialog as fd, messagebox as mBox
#from queue import Queue

import numpy as np
import threading
import time






from arci_ToolTip_refact import ToolTip as tt
from arci_different_tabs_refact import different_tabs 
#from MySQL import MySQL
#from Resources import I18N
#from arci_callbacks_refact import Callbacks
from arci_scheduler_refact import Espec_Scheduler

from espec_refact import Espec_Communication
from position_refact import Aero_Position_Communication
from acquisition_refact import Acquisition
from datalogger_refact import Datalogger_Communication
from analysis_refact import Analysis


# Module level GLOBALS
GLOBAL_CONST = 42

#===================================================================
class Window():
    
    root_folder="D:/Auto/"
    resourcelocation=root_folder+'resources/'
    datalogger_title=u'Configuration - 2 - BenchLink Data Logger 3'
    
    
    
    
    def __init__(self):
        
        
        self.defaultlocation=Window.root_folder
        
        #======================
        # Create instance
        #======================
       
        self.master = tk.ThemedTk()
        

        # Disable resizing the window
        #self.master.resizable(0,0)
        lst_themes=["aquativo", "arc","black","blue","clearlooks","equilux",\
                    "itft1","smog", \
                    "elegance","kroc","keramik",\
                    "plastik","radiance",\
                    "winxpblue"]
        
        self.master.set_theme("itft1")
        #self.master.set_theme(lst_themes[np.random.randint(low=0, high=14)])
        
        
        #(1 #final run),(0 atleastconnect) (-1 very fast checking)
        self.realrun=-1
        
        

        
        #different tabs now in different module
        self.diff_tabs=different_tabs(self)
        
    
        self.especscheduler=Espec_Scheduler(self)
        self.espec_comm=Espec_Communication(self,self.realrun)
        self.aerotec_comm=Aero_Position_Communication(self,self.realrun)
        self.datalogger_comm=Datalogger_Communication(self,self.realrun)
        self.acquisition=Acquisition(self,self.realrun)
        self.analysis_obj=Analysis(self,self.realrun)
        
        
        # Add a title
        self.master.title("DATA ACQUISITION SYSTEM ARDF")
       
        
        
        
        
        #populating the tabs
        self.diff_tabs.createtabs()
        self.especscheduler.populate_tab_schedule_espec()
        self.especscheduler.populate_tab_retrieveschedule()
        self.especscheduler.populate_tab_deleteschedule()
        
        self.diff_tabs.populate_tab_input()
        
        
        def closing_protocol(self):
            pass
        
        
#       
    #####################################################################################
    

if __name__ == '__main__':
    #======================
    # Start GUI
    #======================
    window = Window()
#    print(oop.log)
##     oop.log.setLoggingLevel(oop.level.DEBUG)
#    oop.log.setLoggingLevel(oop.level.MINIMUM)
#    oop.log.writeToLog('Test message')
    window.master.mainloop()
