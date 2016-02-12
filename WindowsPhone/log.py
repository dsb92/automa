#!/usr/bin/python
# -*- coding: utf-8 -*-
from datetime import datetime
import os

current_date = datetime.now().date().isoformat()
submitlog = "log_submit_"+str(current_date)+".txt"
rootdir = os.path.dirname(os.path.realpath(__file__))

def printMessage(msg, waitForUser=False):
    os.chdir(rootdir)
    f = open(submitlog, 'a')
    
    if waitForUser:
        return raw_input ("------------> " + msg)
    else:
        now = datetime.now()
        nowtime = str(now.strftime("%H:%M:%S"))
        f.write(nowtime + "-------" + msg + "\n")
        print nowtime + "-------" + msg
        f.close()
