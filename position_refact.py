# -*- coding: utf-8 -*-
import pathlib
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import sys
import time


class Aero_Position_Communication():
    def __init__(self,window,realrun):
        
        self.window=window
        self.realrun=realrun
        #Validating the Aerotech Position Controller
        self.aerotech_url="http://192.168.1.16"
        #Install the latest Chrome
        path_to_chromedriver=pathlib.Path(self.window.defaultlocation+'chromedriver_win32/chromedriver.exe')
        self.driver=""
        #print(path_to_chromedriver)
        #self.driver=webdriver.Chrome("D:\\Auto\\resources\\chromedriver_win32\\chromedriver.exe")
        
        if self.realrun==1 or self.realrun==0:
            self.driver=webdriver.Chrome(str(path_to_chromedriver))
            print("Waiting for the AEROTECH URL....")
            self.driver.get(self.aerotech_url)
            time.sleep(10)
            #Checking if the two_balls areon the screen
            balls_Element=self.driver.find_element_by_id("enableDisableAxis0")
    
            if balls_Element is None:
                print("The position controller is not connecting...")
                self.window.master.destroy()
                sys.exit(0)
            else:
                print("Connected to Aerotech controller...")
            
            if self.realrun==1:
                #taking care of position homing and enable X , absolute
                self.aerotech_enab_absol()
               
    #This is the function that actually takes care of homing, enable,absolute in the aerotech position
    def aerotech_enab_absol(self):
        #Clicking two balls on the screen
        balls_Element=self.driver.find_element_by_id("enableDisableAxis0")

        if balls_Element is None:
            print("The position controller is not connecting...")
            sys.exit()
        else:
            print("Clicked Balls and waiting 60 seconds... ")
            balls_Element.click()
            time.sleep(1)
            self.driver.implicitly_wait(60)


        #Clicking home button
        print("Clicked home button and waiting 60 seconds... ")
        self.driver.find_element_by_id("homeAxis0").click()
        self.driver.implicitly_wait(60)



        #sending ENABLE X and pressing enter
        imme_comm=self.driver.find_element_by_id("immediate-command-text")
        imme_comm.send_keys("ENABLE X")
        time.sleep(1)
        imme_comm.send_keys(Keys.RETURN)
        time.sleep(5)


        #checking for NO ERROR in the  bottom status bar
        #bottombar_Element=driver.find_element_by_id("status-bar")
        #if bottombar_Element is not None:
        #    print("bottombar_Element found")
        #bottombar_value=bottombar_Element.get_attribute("value")
        #print("The bottombar value is " +bottombar_value)


        check_enable_Element = self.driver.find_element_by_id('axis0Status')
        check_enable_Text = check_enable_Element.text

        if check_enable_Text=='Enabled':
            print("Enabled and lets continue")
            time.sleep(0.5)

        #clearing the immediate-command text for next command and then Absolute command
        print("1")
        imme_comm=self.driver.find_element_by_id("immediate-command-text")
        print("2")
        time.sleep(2)
        print("3")
        imme_comm.clear()
        print("4")
        imme_comm.send_keys("ABSOLUTE")
        print("5")
        time.sleep(1)
        print("6")
        imme_comm.send_keys(Keys.RETURN)
        print("7")
        time.sleep(5)   
    
    
    
    
    
    #Changing the position to the desired value
    def desired_aerotech_pos(self,posnum):
        imme_comm=self.driver.find_element_by_id("immediate-command-text")
        time.sleep(0.5)
        imme_comm.clear()
        if posnum==1:
            imme_comm.send_keys("RAPID X -0.542000 F5")
            desired_pos=-0.542000
        elif posnum==2:
            imme_comm.send_keys("RAPID X 89.458000 F5")
            desired_pos=89.458000
        elif posnum==3:
            imme_comm.send_keys("RAPID X 179.45800 F5")
            desired_pos=179.45800
        elif posnum==4:
            imme_comm.send_keys("RAPID X 269.45800 F5")
            desired_pos=269.45800
        else:
            print("Invalid Position Number...")
            sys.exit(0)

        print("waiting for two minutes to set to desired position")
        time.sleep(1)
        imme_comm.send_keys(Keys.RETURN)
        time.sleep(60)
        #self.driver.implicitly_wait(60)

        #Checking if the position feedback has reached the desired value
        pos_feedback_Text=0

        while abs(desired_pos-float(pos_feedback_Text))>10**-3:
            pos_feedback_Element=self.driver.find_element_by_id('axis0PosFbk')
            time.sleep(5)
            pos_feedback_Text = pos_feedback_Element.text

        print("Reached the desired position")
        return 1
    