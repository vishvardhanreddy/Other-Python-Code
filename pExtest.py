#!/usr/bin/env python

import os
import sys
import datetime
import pexpect
import time
import telnetlib
import re


def main():
    
    pwd = "cisco"
    user = "admin"
    host = "10.29.209.253"
    
    if sys.argv[0]:
        print("All:%s") % ( sys.argv )
        for vals in sys.argv:
            if 'ss' in vals:
                connect_type = via_ssh
            elif 'te' in vals:
                connect_type = via_tel

        if len(sys.argv) > 3:
            pwd = sys.argv[1]
            user = sys.argv[2]
            host = sys.argv[3]
        elif len(sys.argv) > 2:
            pwd = sys.argv[1]
            user = sys.argv[2]
        elif len(sys.argv) > 1:
            pwd = sys.argv[1]
            
        print("Passed:\n\tHost:%s\n\tUser:%s\n\tPwd:%s\n\tConnect by:%s") % ( host, user, pwd, connect_type )

    myShows = run_show(host, user, pwd,connect_type)
    if len(myShows) > 100:
        print("\nSuccess !!!\n")
        
    print("\nReturn: %s") % ( myShows )
    
    

def run_show(host, user, pwd,connect_type):
    command_stack = []
    command_stack = ["show log last 30"]

    pc = pConnect(host, user, pwd,connect_type)
    for cmds in command_stack:
        print("Commands:\n:%s") % ( cmds )

    show_out = pc.run_commands(pc, command_stack, False)

    return(show_out)

class pConnect():
    import smtplib
    import email.utils
    from email.mime.text import MIMEText

    def __init__(self,host,user,pwd,connect_type = 0):
        self.connectStr = self.connectStr(host,user)
        self.debug = 1
        self.logged_in = 0
        self.host = host
        self.user = user
        self.pwd = pwd
        self.time_out = 600
        self.xml_end = "]]>]]>"
        self.netconf = 'netconf format'
        self.sent_netconf = 0
        self.connect_type = connect_type   ### 0 = ssh,1 = Telnet
        if self.debug:
            print("\nINIT()->pConnect() Called\n")
        self.DataDir = "./"            
    
    def connectStr(self, host, user):
        return("".join(["ssh -vv ",user ,"@",host]))

    def ExPyConnect(self):
        log_records = ""
        try:
            pEx9k = pexpect.spawn (self.connectStr, timeout=self.time_out, maxread=4096)    #1048567
        except Exception as e:
            print("ERROR !!!! : %s") % ( str(e) )
            return(e)
    
        pEx9k.expect('[p|P]assword:')
        pEx9k.sendline(self.pwd)
        pEx9k.expect(['#',pexpect.EOF])
        pEx9k.sendline("terminal length 0")
        pEx9k.expect(['#',pexpect.EOF])
        pEx9k.sendline("terminal width 0")
        pEx9k.expect(['#',pexpect.EOF])
        self.logged_in = 1
        return (pEx9k)
    
    def end_netconf(self):
        return("""<?xml version="1.0" encoding="UTF-8" ?>
        <rpc message-id="106" xmlns="urn:ietf:params:xml:ns:netconf:base:1.0">
            <close-session/>
        </rpc>""")

    ### Log out of Device
    def ssh_logout(self):
        if self.logged_in:
            self.logged_in = int(0)
            self.pEx9k.close()

    def run_commands(self, pConn, cmd, toFile = True):
        """
            generates a netconf query to a Cisco IOS-XR router from 
            a provided list of tags for the XML request
            Params: cmd = Command / XML to process
                    toFile = '' | File name to write out results
        """
        #### REd prompt for Telnet connect type 
        REd = re.compile('^[A-Za-z]{1,3}\s[A-Za-z]{1,3}\s+\d{1,2}\s\d{1,2}\:\d{1,2}\:\d{1,2}\.\d{1,3}',re.IGNORECASE)
        rpc_timeout = int(600)
        sleep_time = .5
        date = datetime.datetime.now().strftime('%Y-%m-%d')
        rpcReturn = ""
        command_type = ""
        cType = "C"
        command_list = dict()
        wait_for = ["#",']]>]]>',REd, self.xml_end, pexpect.EOF, pexpect.TIMEOUT]
        msgborder = "\n"+(40 * '-=')+"\n"

        ##### Iterate over inbound list
        ##### Set XML/TXT type for process and return
        ##### Build command_list dict() of Command:Type
        ##### Types : XML ( NetConf ), TXT (Textual), TML (CLI => XML out)

        for c in cmd:
            if self.debug:
                print("Command about to run: %s") % ( c ) 
            if "xml version" in c:
                command_type = "XML"
            elif " xml" in c:
                command_type = "TML"
                self.sent_netconf = 0
            else:
                command_type = "TXT"
                self.sent_netconf = 0

            command_list[c.strip()] = command_type
            if self.debug:
                print("Command Type to run: %s") % ( command_type )

        if self.debug:
            print("Cmd:Type is : %s:%s") % ( c.strip(), command_type )

        if not self.logged_in:
            if self.connect_type == 0:
                print(msgborder+"\nConnect via SSHv2\n"+msgborder)
                pEx9k = self.ExPyConnect()
            else:
                print(msgborder+"\nConnect via Telnet\n"+msgborder)
                pEx9k = self.cli_connect(pConn.host,pConn.user, pConn.pwd)
                promptOff, promptMatch, cmd_result = pEx9k.expect(["#"],15)


        if self.debug:
            print("Setting NetConf enable VAR:%d") % (self.sent_netconf)

        #### Use badCommand to indecate when 'Invalid input detected' in result buffer
        badCommand = 0
        ### use command_list dict() built above ( command : type )
#        for cmds, command_type in command_list.iteritems():
        for cmds in cmd:                           
            if self.debug:
                print("Execute command: %s(type:%s)" ) % ( cmds, command_list[cmds] )
        
            if toFile:
                if not os.path.isdir("./"):
                    try:
                        myDataDir = os.mkdir(self.DataDir)
                    except Exception as e:
                        return(str(e))


            if self.debug:
                print("What CMD is sent: %s (%s)") % ( cmds, command_list[cmds] )
            if self.connect_type == 0:
                pEx9k.sendline(cmds)
            else:
                pEx9k.write(cmds+"\n")

            time.sleep(sleep_time)
            cmd_result = ""
            if self.connect_type == 0:
                print("Before:%s\nAfter%s") % ( str(pEx9k.before), str(pEx9k.after))
                pEx9k.expect(wait_for)
                cmd_result = str(pEx9k.before)                       
            else:
                cmd_result = pEx9k.read_until('#',30)
                print("(3)Result: => %s" ) % ( cmd_result )


            if 'Invalid input detected' in cmd_result:
                badCommand = 1
            else:
                badCommand = 0

        self.logged_in = 0

        pEx9k.close()
        
        return(cmd_result.strip())

    def cli_connect(self,host, user, pwd):
        if not host:
            host = raw_input("Enter a host to begin testing: ")
                
        if not host:
            print("Host IP / DNS Required, save a step, pass -d <IP address> on Command Line\r\n")
            return(False)

        tn = telnetlib.Telnet(host)
        time.sleep(1)
        tn.set_debuglevel(10)
    
        if not user:
            tn.read_until("username: ")
            tn.write(user + "\r\n")
        else:
            tn.write(user + "\n")
    
        if not pwd:
            pwd = getpass.getpass()
            if pwd:
                tn.read_until("password: ")
                tn.write(pwd + "\n")
        else:
            tn.write(pwd + "\n")
    
        tn.write("terminal length 0\n")
        tn.read_until("#",10)
        tn.write("terminal width 0\n")
        tn.read_until("#",10)
        
        return(tn)
    
    def cli_exit(self,asr9k):
        asr9k.write("exit \r\n")
        return(True)
    

if __name__ == '__main__':
    global via_ssh
    global via_tel
    global user
    global pwd
    global host
    via_ssh = 0
    via_tel = 1
    main()
    

