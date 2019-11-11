# -*- coding: utf-8 -*-

import sys
import time
import numpy as np
import pandas as pd

import pathlib
import os, stat
import shutil

class Acquisition():
    def __init__(self,window,realrun):
        
        self.window=window
        self.realrun=realrun
        
    
    def completeacquisition(self):

        self.acq_starts_in_10sec()
        if self.window.diff_tabs.comboday_set == 'L1':
            self.createdirectory(self.window.diff_tabs.base_folder_path)
        
        self.createdirectory(self.window.diff_tabs.combolocation)
        #taking care of position homing and enable X , absolute
        #self.window.aerotec_comm.aerotech_enab_absol()

        if (self.window.diff_tabs.comboday_set == 'L1' or self.window.diff_tabs.comboday_set == 'L2' or self.window.diff_tabs.comboday_set == 'L3'):
            self.window.acquisition.acquistionL123()

        if self.window.diff_tabs.comboday_set == 'D4':
            self.window.acquisition.acquistionD4()

        if self.window.diff_tabs.comboday_set == 'D5':
            self.window.acquisition.acquistionD5()
        
        if self.window.diff_tabs.comboday_set == 'D6':
            self.window.acquisition.endurancetest()
          
    def acquistionL123(self):

        #lsttemp=np.array([70,42.5,15,-12.5,-40,-12.5,15,42.5,70])
        lstramp=np.array([30,15,15,25,35,15,15,15,15])
        lstramp=lstramp[0:self.window.diff_tabs.tempruns_set]
        lstrampdelay=(lstramp+40)*0.01
#        lstrampdelay=(lstramp+40)*0.01
        #range starts from 0
        #for tempiter in range(1,int(self.tempruns_set)+1):
        for tempiter,actual_temp in enumerate(lstrampdelay):
            print("1st sleep")
            tempiter=tempiter+1
            #self.check_temp_reached(actual_temp)
            #time.sleep((lstrampdelay[tempiter-1]+40)*60)
            time.sleep((lstrampdelay[tempiter-1])*60)
            #---position one start--------
            #position numbers are importsnt as desired aerotech position depends on this number
            
            positionnumber=1
            print(f"positionnumber={positionnumber}")
            self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
            self.window.acquisition.autoitdataloggerL123(tempiter,positionnumber)
            
            #-------finished one position------------------------------------------

            positionnumber=2
            print(f"positionnumber={positionnumber}")            
            self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
            self.window.acquisition.autoitdataloggerL123(tempiter,positionnumber)
            
            #-------finished two position------------------------------------------

            positionnumber=3
            print(f"positionnumber={positionnumber}")  
            self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
            self.window.acquisition.autoitdataloggerL123(tempiter,positionnumber)
           

            #-------finished THREE position------------------------------------------

            positionnumber=4
            print(f"positionnumber={positionnumber}")  
            self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
            self.window.acquisition.autoitdataloggerL123(tempiter,positionnumber)
            


    def createdirectory(self,location):
        if not os.path.exists(location):
            try:
                os.mkdir(location)
            except OSError:
                print ("Creation of the directory locaton %s failed " % location)
                self.window.master.destroy()
                sys.exit(0)
            else:
                print ("Successfully created  directory %s " %location)
                
                
    def checkfileexists(self, checkfilename):
        path_checkfilename=pathlib.Path(checkfilename)
        if not path_checkfilename.exists():
            print("Oops"+ checkfilename +" file doesn't exist!\n")
            self.window.master.destroy()
            sys.exit(0)
        else:
            print(checkfilename)
            return True

    def acquistionD4(self):
        #self.tempruns_D45_set  is number of temperature runs set for day 4 and day5
        lsttemp=[-40,70,20]
        #for tempiter in range(1,int(self.tempruns_D45)+1):
        for tempiter,actual_temp in enumerate(lsttemp):

            tempiter=tempiter+1
#            #self.check_temp_reached(actual_temp)
#            #step 1 waiting for temperature ramp for 1 hr
#            time.sleep(60*60)
#            #step 2 waiting at that temperature for 1 hr
#            time.sleep(60*60)


            #step 3 Acquistion by altering the position
            for cycle in range(1,7):
            # when ever  cycle is zero it is going to create a new file in autoitdatalogger45

                positionnumber =1
                self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)

                positionnumber =2
                self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)

#            #step 4 waiting for 30 minutes
#            time.sleep(30*60)

            #step 5 Acquistion by altering the position 3 and 4
            for cycle in range(1,7):
            # when ever  cycle is zero it is going to create a new file in autoitdatalogger45

                positionnumber =3
                self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)

                positionnumber =4
                self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)


#            #step 6 waiting for 30 minutes
#            time.sleep(30*60)


            #step 7 Acquiring 10 readings in position
            #temp_33 shall be the file name where the readings have been taken for 10 times
            positionnumber =33

            self.window.aerotec_comm.desired_aerotech_pos(3)

            for cycle in range(1,11):
                #print ("Start : %s" % time.ctime())
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)
                #print ("Stop : %s" % time.ctime())

        self.summaryD4()



    def finalanalysisD45(self):
        dfD4summary_temp1=pd.read_csv('D4summary_temp1.txt')
        dfD5_temp1_pos5= pd.read_csv('raw_text_D5_1_5.txt')
        dfD5_temp1_pos6= pd.read_csv('raw_text_D5_1_6.txt')

        df_finalD45_temp1=pd.concat([dfD4summary_temp1,dfD5_temp1_pos5,dfD5_temp1_pos6],axis=1)
        df_finalD45_temp1.to_csv('finalD45_temp1.txt',index=False)


        dfD4summary_temp2=pd.read_csv('D4summary_temp2.txt')
        dfD5_temp2_pos5= pd.read_csv('raw_text_D5_2_5.txt')
        dfD5_temp2_pos6= pd.read_csv('raw_text_D5_2_6.txt')

        df_finalD45_temp2=pd.concat([dfD4summary_temp2,dfD5_temp2_pos5,dfD5_temp2_pos6],axis=1)
        df_finalD45_temp2.to_csv('finalD45_temp2.txt',index=False)

        dfD4summary_temp3=pd.read_csv('D4summary_temp3.txt')
        dfD5_temp3_pos5= pd.read_csv('raw_text_D5_3_5.txt')
        dfD5_temp3_pos6= pd.read_csv('raw_text_D5_3_6.txt')

        df_finalD45_temp3=pd.concat([dfD4summary_temp3,dfD5_temp3_pos5,dfD5_temp3_pos6],axis=1)
        df_finalD45_temp3.to_csv('finalD45_temp3.txt',index=False)
        
        
        dfD4_temp1_pos33=pd.read_csv('raw_text_D4_1_33.txt')
        dfD4_temp2_pos33=pd.read_csv('raw_text_D4_2_33.txt')
        dfD4_temp3_pos33=pd.read_csv('raw_text_D4_3_33.txt')



    def summaryD5(self):
        """
        this Concat does along the columns
        """

        if self.tempruns_D45_set==3:

            dfD5_temp1_pos1= pd.read_csv('raw_text_D5_1_1.txt')
            dfD5_temp1_pos2= pd.read_csv('raw_text_D5_1_2.txt')
            
            D5summary_temp1=pd.concat([dfD5_temp1_pos1,dfD5_temp1_pos2],axis=1)
            D5summary_temp1.to_csv('D5summary_temp1.txt',index=False)

            dfD5_temp2_pos1= pd.read_csv('raw_text_D5_1_1.txt')
            dfD5_temp2_pos2= pd.read_csv('raw_text_D5_1_2.txt')
           

            D5summary_temp2=pd.concat([dfD5_temp2_pos1,dfD5_temp2_pos2],axis=1)
            D5summary_temp2.to_csv('D5summary_temp2.txt',index=False)

            dfD5_temp3_pos1= pd.read_csv('raw_text_D5_3_1.txt')
            dfD5_temp3_pos2= pd.read_csv('raw_text_D5_3_2.txt')
            

            D4summary_temp3=pd.concat([dfD5_temp3_pos1,dfD5_temp3_pos2],axis=1)
            D4summary_temp3.to_csv('D5summary_temp3.txt',index=False)



    def summaryD4(self):
        """
        this Concat does along the columns
        """

        if self.tempruns_D45_set==3:

            dfD4_temp1_pos1= pd.read_csv('raw_text_D4_1_1.txt')
            dfD4_temp1_pos2= pd.read_csv('raw_text_D4_1_2.txt')
            dfD4_temp1_pos3= pd.read_csv('raw_text_D4_1_3.txt')
            dfD4_temp1_pos4= pd.read_csv('raw_text_D4_1_4.txt')

            D4summary_temp1=pd.concat([dfD4_temp1_pos1,dfD4_temp1_pos2,dfD4_temp1_pos3,dfD4_temp1_pos4],axis=1)
            D4summary_temp1.to_csv('D4summary_temp1.txt',index=False)

            dfD4_temp2_pos1= pd.read_csv('raw_text_D4_2_1.txt')
            dfD4_temp2_pos2= pd.read_csv('raw_text_D4_2_2.txt')
            dfD4_temp2_pos3= pd.read_csv('raw_text_D4_2_3.txt')
            dfD4_temp2_pos4= pd.read_csv('raw_text_D4_2_4.txt')

            D4summary_temp2=pd.concat([dfD4_temp2_pos1,dfD4_temp2_pos2,dfD4_temp2_pos3,dfD4_temp2_pos4],axis=1)
            D4summary_temp2.to_csv('D4summary_temp2.txt',index=False)

            dfD4_temp3_pos1= pd.read_csv('raw_text_D4_3_1.txt')
            dfD4_temp3_pos2= pd.read_csv('raw_text_D4_3_2.txt')
            dfD4_temp3_pos3= pd.read_csv('raw_text_D4_3_3.txt')
            dfD4_temp3_pos4= pd.read_csv('raw_text_D4_3_4.txt')

            D4summary_temp3=pd.concat([dfD4_temp3_pos1,dfD4_temp3_pos2,dfD4_temp3_pos3,dfD4_temp3_pos4],axis=1)
            D4summary_temp3.to_csv('D4summary_temp3.txt',index=False)

    def autoitdataloggerD45(self,tempiter,positionnumber,cycle):

        self.autoit_singlecycleclicks()

        """
        The methodology used here is for every new temperature and position file is created.
        temp_position 1_1 to 1_4 , 2_1 to 2_4, 3_1 to 3_4
        in cycle 1 the file is created as it is in write mode. When ever heading is printed it is going
        to create file. The last cycle shall be in 33 file.1_33,2_33,raw_text_D4_3_33
        total 15 files shall be outputon day 4.



        """
        files = []
        # r=root, d=directories, f = files
        for r, d, f in os.walk(self.window.diff_tabs.path_to_datalogger_default):
            for file in f:
                files.append(os.path.join(r, file))


        #Assuming it has saved in dataloggerdefault_file.CSV
        #Name to which each of the cycles file shall be saved
        self.D45_path=pathlib.Path(self.window.diff_tabs.base_folder+'raw_text_'+self.window.diff_tabs.comboday_set+'_'+str(tempiter) +'_'+ str(positionnumber)+'.txt')
        with open(files[0],encoding='UTF-16') as f:
            df_datalogger_default = pd.read_csv(f,skiprows=20,header=None)
            if (cycle == 1):
                df_datalogger_default.loc[[0]].to_csv(self.D45_path,index=False, header=False )  #heading
            df_datalogger_default.loc[[1]].to_csv(self.D45_path,index=False, header=False, mode='a')


        my_file=files[0]
        to_file=pathlib.Path(self.window.diff_tabs.combolocation + str(tempiter) + '_' + str(positionnumber) + '_' + str(cycle)+'.txt')
        shutil.move(my_file, to_file)


    def acquistionD5(self):
        #self.tempruns_D45_set  is number of temperature runs set for day 4 and day5
        lsttemp=[-40,70,20]
        #for tempiter in range(1,int(self.tempruns_D45)+1):
        for tempiter,actual_temp in enumerate(lsttemp):

            tempiter=tempiter+1
#            #self.check_temp_reached(actual_temp)
#            #step 1 waiting for temperature ramp for 1 hr
#            time.sleep(60*60)
#            #step 2 waiting at that temperature for 1 hr
#            time.sleep(60*60)

            #step 3 Acquistion by altering the position
            for cycle in range(1,7):
            # when ever  cycle is zero it is going to create a new file in autoitdatalogger45

                positionnumber =1
                self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)

                positionnumber =2
                self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
                self.window.acquisition.autoitdataloggerD45(tempiter,positionnumber,cycle)           
            
        self.summaryD5()            
            
    def acquisitionD6(self):
        positionnumber=1
        self.window.aerotec_comm.desired_aerotech_pos(positionnumber)
        for run in range(1,15):
#            time.sleep(2*60)
            self.window.acquisition.autoitdataloggerL123(run,positionnumber)
           








                
    def acq_starts_in_10sec(self):

        print("Acquisition starts in 10 seconds")
        for i in range(10,0,-1):
            print("{}\n".format(i))
            time.sleep(1)





    #This function gets the data from the dataloger using autoit and using defaultlocation of the dataloger file
    #for 1 cycle. Then it iterates through it.
    def autoitdataloggerL123(self,tempiter,positionnumber):

        self.autoit_singlecycleclicks()
        files = []
        # r=root, d=directories, f = files
        for r, d, f in os.walk(self.window.diff_tabs.path_to_datalogger_default):
            for file in f:
                files.append(os.path.join(r, file))


        with open(files[0],encoding='UTF-16') as f:
            df_datalogger_default = pd.read_csv(f,skiprows=20,header=None)
            if (positionnumber == 1 and tempiter == 1):
                df_datalogger_default.loc[[0]].to_csv(self.window.diff_tabs.raw_text_path,index=False, header=False )  #heading
            df_datalogger_default.loc[[1]].to_csv(self.window.diff_tabs.raw_text_path,index=False, header=False, mode='a')#This is where data is

        #combolocation folder is created when the acquisition is started for L123
        my_file=files[0]
        to_file=pathlib.Path(self.window.diff_tabs.combolocation + str(tempiter) + '_' + str(positionnumber) +'.txt')
        shutil.move(my_file, to_file)

    def autoit_singlecycleclicks(self):


        #Emptying all the files from the default directory of Datalogger saver so that only one file is created by datalogger3
        files = []
        # r=root, d=directories, f = files
        for r, d, f in os.walk(self.window.diff_tabs.path_to_datalogger_default):
            for file in f:
                files.append(os.path.join(r, file))
        #print(files)

        for i in files:
            os.remove(i)

        print('Removed all files from datalogger default..starting acquisition')
        self.window.datalogger_comm.pywinautosequence()
        