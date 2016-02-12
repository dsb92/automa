#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import log as log
import shutil, sys
from automa.api import *

appIcon = "image_99x99.png"
smallTile = "image_202x202.png"
medTile = "image_300x300.png"
productImage = "image_300x300.png"
splashScreenImage = "image_768x1280.png"
wp8Path = r"/Windows/WP8"
graphicFolder = r"/ready_v13/ready"
screenshotFolder = "ScreenshotWP_"

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
                    graphicPath = scriptdir+graphicFolder + "/" + folder + wp8Path
                    log.printMessage("Returning graphic path " + graphicPath)
                    return graphicPath
                else:
                    #Prøv igen, men kig kun på starten af grafik titel. Samme som slash eksempel men her indeholder grafiktitel ikke slash
                    graphicTitel = graphicTitel.split("-",1)[0]
                    if graphicTitel in folder:
                        log.printMessage("Graphic titel found, replaced with ''...")
                        graphicPath = scriptdir+graphicFolder + "/" + folder + wp8Path
                        log.printMessage("Returning graphic path " + graphicPath)
                        return graphicPath
    except:
        log.printMessage("Failed in graphic path function...")


def getScreenshotPath(root):
    path = root+"/"+screenshotFolder
    log.printMessage("Returning screenshot path " + path)
    return path

# FOR TEST
def loopAllGraphics():
    rootdir = r"C:\Users\David\Projekter\aMediaTest\Device\NytSetup\WindowsPhone"
    temp = "appnames.txt"    
    userhome = os.path.expanduser('~')
    desktop = userhome + '/Desktop/'
    os.chdir(desktop)
    tempLocation = os.getcwd()+"/"+temp

    datafile = file(tempLocation)
    for line in datafile:
        log.printMessage(str(getGraphicPath(rootdir, line)))
