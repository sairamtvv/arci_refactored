# -*- coding: utf-8 -*-
import pywinauto
from pywinauto.application import Application
from pywinauto.findwindows import WindowAmbiguousError, WindowNotFoundError
from pywinauto.controls.common_controls import TabControlWrapper
from pywinauto.keyboard import send_keys, KeySequenceError
import time


class Datalogger_Communication():
    def __init__(self,window,realrun):
        
        self.window=window
        self.realrun=realrun
        
        
        if self.realrun==1 or self.realrun==0:
            self.pywinautosequence()
            
            
    def pywinautosequence(self):
        try:
            app = Application(backend="win32").connect(title=u'Configuration - 2 - BenchLink Data Logger 3', class_name='WindowsForms10.Window.8.app.0.33c0d9d')
            main_dlg = app[u'WindowsForms10.Window.8.app.0.33c0d9d']
            main_dlg.wait('visible')
            print("1")
            time.sleep(1)
            main_dlg.set_focus()
            print("2")
            
            p=main_dlg.TabControl.select(u'Scan and Log Data')
            #print(p.get_properties())
            
            p.client_rects()
            
            scanlog_dlg=main_dlg[u'WindowsForms10.Window.8.app.0.33c0d9d7']
            #scanlog_dlg.child_window(auto_id="m_gridInst", control_type="C1.Win.C1FlexGrid.C1FlexGrid").print_control_identifiers()
            scanlog_dlg.draw_outline()
            scanlog_dlg.click()
            time.sleep(0.5)
            
            
            
            send_keys('^{RIGHT}{DOWN}{LEFT}{LEFT}{ENTER}')
            
            
            app2 = Application().connect(title=u'Set Data Log Fields', class_name='WindowsForms10.Window.8.app.0.33c0d9d')
            setdatalogfield_dlg = app2[u'WindowsForms10.Window.8.app.0.33c0d9d']
            print("11")
            setdatalogfield_dlg.wait('visible')
            print("22")
            
            checkboxbutton = setdatalogfield_dlg.Button3
            checkboxbutton.click()
            
            okbutton=setdatalogfield_dlg.OK
            okbutton.click()
            
            time.sleep(1)
            main_dlg.set_focus()
            send_keys('{RIGHT}{ENTER}')
            
            app3 = Application().connect(title=u'Scan and Log Data Summary', class_name='WindowsForms10.Window.8.app.0.33c0d9d',timeout=150)
            print("50")
            scananddata_dlg = app3[u'Scan and Log Data Summary']
            scananddata_dlg.wait('visible',2*60,5).close()
            
        except(WindowNotFoundError):
            print ('window not found')
            
        except(WindowAmbiguousError):
            print ('There are too many  windows found')
             
        
    def dataloggerfor33_10conseq(self):
        try:
            app = Application(backend="win32").connect(title=u'Configuration - 2 - BenchLink Data Logger 3', class_name='WindowsForms10.Window.8.app.0.33c0d9d')
            main_dlg = app[u'WindowsForms10.Window.8.app.0.33c0d9d']
            main_dlg.wait('visible')
            print("1")
            time.sleep(1)
            main_dlg.set_focus()
            print("2")
            
            p=main_dlg.TabControl.select(u'Scan and Log Data')
            #print(p.get_properties())
            
            p.client_rects()
            
            scanlog_dlg=main_dlg[u'WindowsForms10.Window.8.app.0.33c0d9d7']
            #scanlog_dlg.child_window(auto_id="m_gridInst", control_type="C1.Win.C1FlexGrid.C1FlexGrid").print_control_identifiers()
            scanlog_dlg.draw_outline()
            scanlog_dlg.click()
            time.sleep(0.5)
            
            
            
            send_keys('^{RIGHT}{DOWN}{LEFT}{LEFT}{ENTER}')
            
            for num_runs in range(1,11):
                app2 = Application().connect(title=u'Set Data Log Fields', class_name='WindowsForms10.Window.8.app.0.33c0d9d')
                setdatalogfield_dlg = app2[u'WindowsForms10.Window.8.app.0.33c0d9d']
                
                setdatalogfield_dlg.wait('visible')
                
                
                checkboxbutton = setdatalogfield_dlg.Button3
                checkboxbutton.click()
                
                okbutton=setdatalogfield_dlg.OK
                okbutton.click()
                
                time.sleep(0.5)
                main_dlg.set_focus()
                send_keys('{RIGHT}{ENTER}')
                
                app3 = Application().connect(title=u'Scan and Log Data Summary', class_name='WindowsForms10.Window.8.app.0.33c0d9d',timeout=150)
                scananddata_dlg = app3[u'Scan and Log Data Summary']
                scananddata_dlg.wait('visible',2*60,5).close()
                send_keys('{LEFT}{ENTER}')
                
                
            
        except(WindowNotFoundError):
            print ('window not found')
            
        except(WindowAmbiguousError):
            print ('There are too many  windows found')     
        