#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import log as log
import shutil, sys
from automa.api import *

storeLogo = "image_50x50.png"
logo = "image_150x150.png"
smallLogo = "image_30x30.png"
wideLogo = "image_310x150.png"
splashScreen = "image_620x300.png"
w8Path = r"/Windows/W8"
graphicFolder = r"/ready_v13/ready"
screenshotFolder = "ScreenshotWT"

#rootdir = os.path.dirname(os.path.realpath(__file__))

def getGraphicPath(scriptdir, appname):
    try:
        log.printMessage("Getting graphic path...")
        
        graphicTitel = appname

        if graphicTitel is None:
            log.printMessage("App name is NULL")
            graphicTitel = appname
        else:
            log.printMessage("App name is OK")
            if " Digital Utgave" in graphicTitel:
                log.printMessage("App name contains Digital Utgave...replacing it with ''")
                graphicTitel = graphicTitel.replace(" Digital Utgave", "")

            if " " in graphicTitel:
                log.printMessage("App name contains empty space...replacing it with '-'")
                graphicTitel = graphicTitel.replace(" ", "-") 

            #Hvis grafiktitel indeholder slash, så fjern alt før og tjek på det efterfølgende.
            if "/" in graphicTitel:
                log.printMessage("App name contains slash, replacing it with ''...")
                graphicTitel = graphicTitel.replace(graphicTitel.split("/",1)[0], "")
                graphicTitel = graphicTitel.replace("/", "")            

            graphicTitel = graphicTitel.strip()
            log.printMessage("Looking for graphicTitel: " + graphicTitel)
            for folder in os.listdir(scriptdir+graphicFolder):
                if graphicTitel in folder:
                    log.printMessage("Graphic titel found!")
                    graphicPath = scriptdir+graphicFolder + "/" + folder + w8Path
                    log.printMessage("Returning graphic path " + graphicPath)
                    return graphicPath
                else:
                    #Prøv igen, men kig kun på starten af grafik titel. Samme som slash eksempel men her indeholder grafiktitel ikke slash
                    graphicTitel = graphicTitel.split("-",1)[0]
                    if graphicTitel in folder:
                        log.printMessage("Graphic titel found, replaced with ''...")
                        graphicPath = scriptdir+graphicFolder + "/" + folder + w8Path
                        log.printMessage("Returning graphic path " + graphicPath)
                        return graphicPath
    except:
        log.printMessage("Failed in graphic path function...")


# FOR TEST
def loopAllGraphics():
    rootdir = r"C:\Users\David\Projekter\Device\Windows\Windows8"
    temp = "appnames.txt"
    """
    userhome = os.path.expanduser('~')
    desktop = userhome + '/Desktop/'
    os.chdir(desktop)
    """
    os.chdir(rootdir)
    #tempLocation = os.getcwd()+"/"+temp

    datafile = file(temp)
    for line in datafile:
        log.printMessage(str(getGraphicPath(rootdir, line)))
