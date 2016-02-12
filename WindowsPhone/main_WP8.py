#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import log as log

"Setup - Configurable parameters"
googledocs = "Amedia PID WP8"
prefixtitelcolumn = "Visiolink prefix"
urldocs = "https://docs.google.com/spreadsheets/d/1WhbzQ2XvqOo_6R-twjdm7Z8FT1AGrq9QHe-7Ze3VwYk/edit#gid=494592811"
dashboardWindow = "Windows Phone Dev Center"
tempfileName = "temp.txt"
tempfileName2 = "temp2.txt"
xapfolder = "/Bin/ARM/Release"
maxAmountKeywords = 5
sleepTime = 60
#Tab nr. til hvor google docs skal starte til næste run.
docsresetnum = 15
#docsresetnum = 3

InAppAlias = "Enkelt kjøp"
InAppProductIdentifier = "PurchaseSingleDay"
InAppTag = "Amedia"

"Directory from where this script is located"
rootdir = os.path.dirname(os.path.realpath(__file__))
os.chdir(rootdir)

global debugCounter
debugCounter = 13

global startUpFlag
startUpFlag = True

global endFlag
endFlag = False

def startUserInput():
    global submittype
    submittype = log.printMessage("Choose submit type: 'b' for beta or 'l' for live\n", True)

    if submittype == 'b':
        c = log.printMessage("How many beta testers? (Type a number)\n", True)
        i = 1
        failureString = "Write E-mails without non-ascii characthers like æ,ø,å"
        log.printMessage(failureString.decode("utf-8").encode("iso-8859-1"))
        global participants
        participants = []
        while(i <= int(c)):
            parti = log.printMessage(str(i)+": ", True)
            participants.append(parti)
            i += 1
    
    global xappssubmits
    xappssubmits = int(log.printMessage("Type how many apps you wish to upload. (Type a number)\n", True))
    
    log.printMessage("Starting autoupload script...")


if __name__ == "__main__":
    startUserInput()
    execfile('savescreenshotWP.py')
    sleep(5)
    execfile('autosubmitWP.py')
    sleep(5)
    execfile('uploadimages.py')
    sleep(5)


