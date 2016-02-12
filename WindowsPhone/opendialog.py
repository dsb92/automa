# -*- coding: utf-8 -*-
from automa.api import *
import log as log
from win32api import GetSystemMetrics

def getMonitorResolution():
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)
    return Point(x=width, y=height)

def click_upload_all(): 
    if Text("Specify keywords").exists() == False:
        log.printMessage("Getting HARDCOREDED coord for Image button..")
        if getMonitorResolution() == Point(x=1366, y=768):
            click(Point(x=235, y=170))
        elif getMonitorResolution() == Point(x=1920, y=1080):
            click(Point(x=515, y=492))
            
        log.printMessage("HARDCODED - TEST")
        sleep(2)
    else:
        log.printMessage("Getting coord for Image button..")
        click(Point(x=Text("Specify keywords").center.x, y=Text("Specify keywords").y+325))
        log.printMessage("COORD - TEST")
        sleep(2)
        

    log.printMessage("Clicked Upload all")

def open_image_dialog():     
    if Window('Åbn').exists():
        switch_to(Window('Åbn'))
        print "tid plac true"
        click(Button("Tidligere placeringer").center - (10,0))
        return True
    else:
        return False
        

log.printMessage("Uploading all images...")
switchTo(dashboardWindow)
click_upload_all()

i=0
while(open_image_dialog() == False and i<3):
    click_upload_all()
    i+=1

if Window('Åbn').exists() == False:
    raise Exception("ERROR: Could not find dialog to upload images...")
    
