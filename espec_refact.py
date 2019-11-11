# -*- coding: utf-8 -*-

from SCPI_socket import SCPI_sock_connect,SCPI_sock_send
from SCPI_socket import SCPI_sock_close,getDataFromSocket
import time
import sys


class Espec_Communication():
    def __init__(self,window,realrun):
        
        self.window=window
        #espec details i.e temperature controller
        self.espec_HOST="192.168.1.18"
        self.espec_PORT=57732
        self.realrun=realrun
        self.espec_session=""
        if self.realrun==1 or self.realrun==0:
            self.espec_session=SCPI_sock_connect(self.espec_HOST,self.espec_PORT)
            print(self.espec_session)
            if self.realrun==1:
                self.validate_espec()
        
        
        
    def validate_espec(self):    
        
        #Validating ESPEC temperature controller by reading its temperature once
        SCPI_sock_send(self.espec_session,'TEMP?')
        time.sleep(0.5)
        output=self.espec_session.recv(20).decode()
        lstoutput=output.split(",")

        #21.3,-40.0,165.0,-70
        if -110<float(lstoutput[0])<110:
            print("Espec Validated and could read temperature")
        else:
            print("Please Check your connection to ESPEC")
            self.window.master.destroy()
            sys.exit(0)
            
    def check_temp_reached_espec(self,desired_temp):
        SCPI_sock_send(self.espec_session,'TEMP?')
        output=self.espec_session.recv(20).decode()
        lstoutput=output.split(",")
        #21.3,-40.0,165.0,-70
        print(f"Waiting for temperature: {desired_temp}")
        while(float(lstoutput[0])!=desired_temp):
            time.sleep(60)
        print(f"Reached the desired temperature:{desired_temp}")
        
        
    def put_on_espec(self):
        SCPI_sock_send(self.espec_session,'POWER, ON')
        print(self.espec_session.recv(20).decode())
        time.sleep(2)
        
    
    def put_off_espec(self):
        SCPI_sock_send(self.espec_session,'POWER, OFF')
        print(self.espec_session.recv(20).decode())
        time.sleep(2)
        
        
    def set_temp_espec(self,temp):
        #str(temp).format()
        #reqstring=S+temp+ H100.0 L-40.0
        temp=float(temp)
        set_temp="TEMP, S{0:.1f} H100 L-40.0".format(temp)
        SCPI_sock_send(self.espec_session,set_temp)
        print(self.espec_session.recv(20).decode())
        time.sleep(2)
        
        SCPI_sock_send(self.espec_session,'MODE, CONSTANT')
        print(self.espec_session.recv(20).decode())        
    
    def closing_protocol_espec(self):
        SCPI_sock_close(self.espec_session)