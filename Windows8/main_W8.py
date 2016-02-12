#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import log as log

"Setup - Configurable parameters"
googledocs = "Amedia PID W8"
prefixtitelcolumn = "Visiolink prefix"
urldocs = "https://docs.google.com/spreadsheets/d/1e0sswHJo91jPbLfAdzTrC2dfFDbpGP6m8f81aMtP_2A/edit#gid=614091273"
dashboardWindow = "Udviklingscenter - Windows Store-apps"
sendEnAppWindow = "Send en app"
tempfileName = "temp.txt"
tempfileName2 = "temp2.txt"
screenshotFolder = r"/ScreenshotWT"
imageCaption = "skjermbilde"
messageToTestere = "Login: 00009996/test"
packageExt = ".appxupload"
publisherid = "CN=5860D604-48AB-492F-9714-6351AE57306E"

maxAmountKeywords = 7
appSubmits = 70
maxImageUpload = 5
sleepTime = 60
docsresetnum = 19

# Afhængig af hvor mange screenshot, der skal uploades
tab_nr_to_keywords = 14

InAppAlias = "Enkelt kjøp"
InAppProductIdentifier = "PurchaseSingleDay"

"Directory from where this script is located"
rootdir = os.path.dirname(os.path.realpath(__file__))
os.chdir(rootdir)

# OBS ( KUN I DEBUG ) - Skal tælles 1 op for hvert submit.
global debugCounter
debugCounter = 9

global startUpFlag
startUpFlag = True

def startUserInput():
    global xappssubmits
    xappssubmits = int(log.printMessage("Type how many apps you wish to upload. (Type a number)\n", True))
    
    log.printMessage("Starting autoupload script...")


if __name__ == "__main__":
    startUserInput()
    #execfile('savescreenshotW8.py')
    #sleep(5)
    execfile('autosubmitW8.py')



