# -*- coding: utf-8 -*-
from automa.api import *
from time import sleep
import os, sys
import xmlparser
from datetime import datetime
from parameters import *

#Automa setup
Config.search_timeout_secs = 60
Emulator = "Emulator WXGA"
folderName = "ScreenshotWP_"
current_date = datetime.now().date().isoformat()
wpsslog = "log_screenshotWP_"+str(current_date)+".txt"
urldocs = "https://docs.google.com/spreadsheets/d/1WhbzQ2XvqOo_6R-twjdm7Z8FT1AGrq9QHe-7Ze3VwYk/edit#gid=494592811"
screenshotMax = 6
#Directory from where this script is located
rootdir = os.path.dirname(os.path.realpath(__file__))


"****COORDINATION****"

#OBS OBS, koordinatsæt for tasteturet i emulator: (I tilfælde af at visiolink kode ændrer sig, til når man skal logge ind)
a = Point(x=57, y=477)
b = Point(x=234, y=533)
c = Point(x=162, y=533)
d = Point(x=130, y=477)
e = Point(x=113, y=418)
f = Point(x=164, y=477)
g = Point(x=199, y=477)
h = Point(x=234, y=477)
i = Point(x=287, y=418)
j = Point(x=269, y=477)
k = Point(x=303, y=477)
l = Point(x=338, y=477)
m = Point(x=304, y=533)
n = Point(x=269, y=533)
o = Point(x=323, y=418)
p = Point(x=355, y=418)
q = Point(x=45, y=418)
r = Point(x=148, y=418)
s = Point(x=97, y=477)
t = Point(x=185, y=418)
u = Point(x=251, y=418)
v = Point(x=199, y=533)
w = Point(x=79, y=418)
x = Point(x=131, y=533)
y = Point(x=217, y=418)
z = Point(x=96, y=533)

and123 = Point(x=51, y=595)
abcd = Point(x=51, y=595)

num1 = Point(x=45, y=418)
num2 = Point(x=76, y=418)
num3 = Point(x=109, y=418)
num4 = Point(x=146, y=418)
num5 = Point(x=182, y=418)
num6 = Point(x=216, y=418)
num7 = Point(x=249, y=418)
num8 = Point(x=287, y=418)
num9 = Point(x=321, y=418)
num0 = Point(x=354, y=418)

#Capture screenshot
btnAdditionalTools = Point(x=415, y=165)
btnScreenshot = Point(x=687, y=27)
btnCapture = Point(x=500, y=67)
btnSave = Point(x=590, y=67)

#Login 
btnThreeDotsPanel = Point(x=346, y=616)
btnAnnuller = Point(x=290, y=583)
btnLoginFromDotsPanel = Point(x=229, y=420)
tfuser = Point(x=196, y=218)
tfPassword = Point(x=199, y=269)

btnEnter = Point(x=350, y=596)
btnlogin = Point(x=200, y=505)

#Navigation
btnDeviceBack = Point(x=67, y=661)
btnBackInApp = Point(x=57, y=105)

#Zoom in/out from Article View
btnZoomIn = Point(x=340, y=102)
btnZoomOut = Point(x=285, y=104)

#Button from dots panel
btnMyPapersFromDotsPanel = Point(x=169, y=418)
btnSearchFromDotsPanel = Point(x=109, y=418)
btnAboutFromDotsPanel = Point(x=62, y=536)
btnPrivatlivsPolitik = Point(x=106, y=487)
btnSectionOverview = Point(x=200, y=581)

"****END OF COORDINATION****"

def printMessage(msg, waitForUser=False):
    global f
    os.chdir(rootdir)
    f = open(wpsslog, 'a')
    
    if waitForUser:
        return raw_input ("------------> " + msg)
    else:
        now = datetime.now()
        nowtime = str(now.strftime("%H:%M:%S"))
        f.write(nowtime + "-------" + msg + "\n")
        print nowtime + "-------" + msg
        f.close()

def SETAUTO(x):
    Config.auto_wait_enabled = x

def typeNum(num, whattype):
    printMessage("Tabbing, pressing down og up " + str(num) + " times...")
    c = 0
    while (c < num):
        press(whattype)
        c += 1

#Kun vindue/hjemmeside skal maksimeres, når der skiftes.
def switchTo(window):
    switch_to(window)
    
    press(LWIN+UP)

def checkImageExists():
    printMessage("Checking if image exists...")
    if Window("Gem som").exists():
        printMessage("Replacing image...")
        if not Button("Ja").exists():
            press(ENTER)
        else:
            click("Ja")


def changePathToDesktop():
    userhome = os.path.expanduser('~')
    desktop = userhome + '/Desktop/'
    os.chdir(desktop)

def findTempLocation(temp):
    changePathToDesktop()
    global tempLocation
    tempLocation = os.getcwd()+"/"+temp
    

def saveAs(tempfile):
    if Window("Gem som").exists():
        click("Skrivebord")
        write(tempfile, into="Filnavn:")
        click("Gem")
            
    if Window("Bekræft Gem som").exists():
        click("Ja")

def addPrefixesToList():
    printMessage("Adding prefixes from google docs to list...")
    if startUpFlag == True:
        if Window(googledocs).exists():
            switchTo(googledocs)
        else:
            global urlbrowser
            urlbrowser = start("Google Chrome", '-new-window', urldocs)
            sleep(10)
            switchTo(googledocs)
    else:
        switchTo(googledocs)

    press(CTRL+SPACE)
    press(CTRL+'c')

    if startUpFlag == True:
        if Window(tempfileName).exists():
            switch_to(tempfileName)
        else:
            global notepad_1
            notepad_1 = start("notepad")
    else:
        switch_to(tempfileName)

    press(CTRL+'v')
    press(CTRL+'s')
        
    saveAs(tempfileName)
    findTempLocation(tempfileName)

    global prefixes
    prefixes = []
    
    datafile = file(tempfileName)
    for line in datafile:
        prefixes.append(line.strip())

    switch_to(tempfileName)
    press(CTRL+'a',DEL)
    press(CTRL+'s')


def addIfSubmitsToList():
    printMessage("Adding submit status from google docs to list...")

    press(CTRL+SPACE)
    press(CTRL+'c')
    press(RIGHT)

    if startUpFlag == True:
        if Window(tempfileName2).exists():
            switch_to(tempfileName2)
        else:
            global notepad_2
            notepad_2 = start("notepad")
    else:
        switch_to(tempfileName2)
    
    press(CTRL+'v')
    press(CTRL+'s')

    saveAs(tempfileName2)
    findTempLocation(tempfileName2)

    global ifsubmits
    ifsubmits = []
    
    datafile = file(tempfileName2)
    for line in datafile:
        ifsubmits.append(line.strip())

    switch_to(tempfileName2)
    press(CTRL+'a',DEL)
    press(CTRL+'s')

def combinePrefixWithIsSubmit():
    print "tis"
    addPrefixesToList()
    switchTo(googledocs)
    press(RIGHT)
    addIfSubmitsToList()
   
def submissionList():
    printMessage("Combining prefix list with submit status list...")
    i=0
    comb = []
    while(i<len(prefixes)):
        comb.append(str(prefixes[i])+"_"+str(ifsubmits[i]))
        i+=1

    return comb

def openSolutionFile(root, name):
    try:
        click("FILE"), hover("Open"), click("Project/Solution...")
    except:
        sleep(10)
        press(CTRL+SHIFT+'o')

    printMessage("Opening solution file...")
    if Button("Tidligere placeringer").exists() == False:
        press(TAB,TAB,TAB,TAB)
        press(ENTER)
    else:
        click(Button("Tidligere placeringer").center - (10,0))

    write(root)
    press(ENTER)
    #doubleclick(name)
    write(name, into="Filnavn:")
    press(ENTER)
    sleep(sleepTime/2)

def buildSolution():
    press(F6)
    printMessage("Building solution file...")
    #wait_until(Text("Build succeeded").exists, timeout_secs=30, interval_secs=3.0)
    sleep(sleepTime/2)
    printMessage("Build succeeded")

def setupEmulator():
    printMessage("Setting up Emulator...")
    click("Solution Platforms")
    press(DOWN,ENTER)
    click("Solution Configuration")
    press(DOWN,ENTER)
    click(Text("Browser Link").center-(23,0))
    hover(Emulator)
    click(Emulator)

def startWithoutDebugging():
    printMessage("Starting without debugging...")
    press(CTRL+F5)
    sleep(sleepTime)

"""
def checkDeploymentError():
    if not Text("Deploy succeeded").exists():
        if Button("No").exists():
            click("No")
            startWithoutDebugging()
            
    printMessage("Deployment success")
"""

def setupScreenshotFolder(root):
    switch_to(Emulator)
    printMessage("Setting up screenshot folder...")
    screenshotFolder = folderName+os.path.basename(root)
    
    #Setup save location.
    #Capture KioskPage
    click(btnAdditionalTools)
    sleep(2)
    click(btnScreenshot)
    click(btnCapture)
    click(btnSave)

    try:
        wait_until(Window("Gem som").exists, timeout_secs=5)
        switch_to("Gem som")
    except:
        sleep(5)
        switch_to("Gem som")
        
    click(Button("Tidligere placeringer").center - (10,0)), write(root), press(ENTER)
    click("Ny mappe"), write(screenshotFolder), press(ENTER)
    screenshotfolderExists(screenshotFolder)
    global screenshotCount
    screenshotCount+=1
    printMessage("Capturing screenshot " + str(screenshotCount))
    write("Screenshot0"+str(screenshotCount), into="Filnavn")
    click("Gem")

    checkImageExists()
    

def startUpConfigEmulator():
    switch_to(Emulator)
    printMessage("Dragging emulator to (0,0) position...")
    drag(Window(Emulator).left, Point(x=0, y=Window(Emulator).left.y))
    drag(Window(Emulator).top, Point(x=Window(Emulator).top.x, y=0))
    zoomButton45p = Point(x=Window(Emulator).x+Window(Emulator).width+15, y=Window(Emulator).y+135)

    click(zoomButton45p)
    sleep(5)
    try:
        switch_to("Zoom")
        printMessage("Switching to Zoom")
        write("45", into="Percent"), click("OK")
    except:
        printMessage("Tabbing to Zoom")
        press(TAB)
        press(ALT+'a')
        write("45")
        press(ENTER)
        
        
    # Max resolution (2560x1440) for current monitor.. but does not matter.
    drag(Window(Emulator).right, Point(x=2560, y=Window(Emulator).right.y)) 
    drag(Window(Emulator).left, Point(x=0, y=Window(Emulator).left.y))
    drag(Window(Emulator).top, Point(x=Window(Emulator).top.x, y=0))

def screenshotfolderExists(folder):
    if Window("Bekræft erstatning af mappe").exists():
        click("Ja")
        doubleclick(folder)
    else:
        press(ENTER)


def loginOnPaperApp():
    printMessage("Logging into app")
    click(btnThreeDotsPanel)
    click(btnLoginFromDotsPanel)
    click(tfuser)
    Config.auto_wait_enabled = False
    # HER SKAL ÆNDRES, HVIS VISIOLINK BRUGER ÆNDRER SIG/ELLER KUNDENS
    click(and123)
    click(num0, num0, num0, num0, num9, num9, num9, num6)
    click(tfPassword)
    # HER SKAL ÆNDRES, HVIS VISIOLINK KODE ÆNDRER SIG/ ELLER KUNDENS
    click(t,e,s,t)
    Config.auto_wait_enabled = True
    click(Window(Emulator).center)
    click(btnlogin)
    sleep(5)

def captureScreenshot():
    if startUpFlag == True:
        click(btnAdditionalTools)
        click(btnScreenshot)
        click(btnCapture)
        click(btnSave)
    else:
        click(btnCapture)
        click(btnSave)

    try:
        wait_until(Window("Gem som").exists, timeout_secs=5)
        switch_to("Gem som")
    except:
        sleep(5)
        switch_to("Gem som")

    global screenshotCount
    screenshotCount+=1
    printMessage("Capturing screenshot " + str(screenshotCount))

    try:
        write("Screenshot0"+str(screenshotCount), into="Filnavn:")
        click("Gem")
    except:
        #Placering er ikke tilgængelig dialog
        press(ALT+F4)
        write("Screenshot0"+str(screenshotCount), into="Filnavn:")
        click("Gem")
        
    checkImageExists()
    

def beginCapturing(prefix):
    printMessage("Ready!")
    
    Config.auto_wait_enabled = True
    nextPage = Window(Emulator).right - (30,0)
    previousPage = Window(Emulator).left + (30,0)

    click(Window(Emulator).center)
    sleep(8)
    click(nextPage)
    sleep(8)
    #Capture SpreadPage
    printMessage("Capturing SpreadPage...")
    captureScreenshot()
    
    xml = xmlparser.xmlParse_server_xml_todayspaper(prefix)

    #Hvis avis har sections og article view
    if xmlparser.xmlParse_server_has_sections(xml) and xmlparser.xmlParse_server_has_articles(xml):
        printMessage(prefix + " has sections and articles")
        btnArticleView = Point(x=140, y=580)

    #Hvis avis kun har article view (og pages)
    elif xmlparser.xmlParse_server_has_articles(xml):
        printMessage(prefix + " has articles only")
        btnArticleView = Point(x=170, y=580)

    #Hvis avis slet ikke har article view, tag ikke screenshot
    else:
        printMessage(prefix + " has NOT sections or articles")
        btnArticleView = None

    if btnArticleView is not None:
        click(btnThreeDotsPanel)
        click(btnArticleView)
        sleep(8)
        click(Window(Emulator).center)
        sleep(8)
        #Capture from Article View
        printMessage("Capturing Article View...")
        captureScreenshot()
        #Capture list from Article View
        click(btnBackInApp)
        printMessage("Capturing Article List from Article View...")
        captureScreenshot()
        click(btnDeviceBack)


    if xmlparser.xmlParse_server_has_sections(xml) and xmlparser.xmlParse_server_has_articles(xml):
        btnPageOverview = Point(x=260, y=580)

    elif xmlparser.xmlParse_server_has_articles(xml):
        btnPageOverview = Point(x=230, y=580)

    else:
        btnPageOverview = Point(x=200, y=580)

    click(btnThreeDotsPanel)
    click(btnPageOverview)
    sleep(8)
    #Capture Page overview
    printMessage("Capturing Page Overview...")
    captureScreenshot()

    click(btnDeviceBack)
    sleep(8)
    click(btnDeviceBack)
    click(btnThreeDotsPanel)
    click(btnSearchFromDotsPanel)
    #Capture Search function from coverflow
    printMessage("Capturing search function...")
    captureScreenshot()

    """
    #Capture section overview 7
    click(btnDeviceBack)
    sleep(5)
    click(btnThreeDotsPanel)
    click(btnSectionOverview)
    printMessage("Capturing Section Overview...")
    captureScreenshot()
    """

def restart():
    global startUpFlag
    startUpFlag = False

    if Window("Placeringen er ikke tilgængelig").exists():
        switch_to("Placeringen er ikke tilgængelig")
        press(ALT+F4)
    
    if Window("Gem som").exists():
        switch_to("Gem som")
        press(ALT+F4)

    if Window("Zoom").exists():
        switch_to("Zoom")
        press(ALT+F4)

    if Window("Microsoft Visual Studio").exists():
        sleep(3)
        switchTo("Microsoft Visual Studio")


def close():
    if Window(Emulator).exists():
        switch_to(Emulator)
        press(ALT+F4)

    if Window("Emulator").exists():
        switch_to("Emulator")
        press(ALT+F4)

    if Window("Microsoft Visual Studio").exists():
        kill("Microsoft Visual Studio")
        sleep(3)
    

def mainflow(root,name,prefix):
    try:
        printMessage("*****************APP SCREENSHOT NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        if startUpFlag == True:
            VS2013 = start("Visual Studio 2013")
            
        openSolutionFile(root,name)
        buildSolution()
        setupEmulator()
        startWithoutDebugging()
        #checkDeploymentError()
            
        if startUpFlag == True:
            startUpConfigEmulator()

        global screenshotCount
        screenshotCount = 0

        setupScreenshotFolder(root)

        #loginOnPaperApp()
            
        beginCapturing(prefix)

        global startUpFlag
        startUpFlag = False
        printMessage("DONE - Finish capturing for: " + name)
        printMessage("*****************END OF APP SCREENHOT NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appCounter
        appCounter += 1
        switchTo(os.path.splitext(name)[0] + " - Microsoft Visual Studio")
    except:
        print "Unexpected error:", sys.exc_info()[1]
        printMessage("ERROR - Failed capturing for " + name)
        restart()
        


def foreach(prefix):
    for root, dirs, files in os.walk(rootdir+"/"+prefix):  
        for name in files:
            if name.endswith(".sln") and not any(folderName in f for f in os.listdir(root)):
                mainflow(root,name,prefix)
                #print root
                return 0
            elif name.endswith(".sln") and any(folderName in f for f in os.listdir(root)):
                return 0
                
def screenshot(prefix):
    for dir in os.listdir(rootdir):
        prefix = prefix.replace("_WP8", "")
        if prefix == dir:
            foreach(prefix)
            return 0


def runList(slist):
    for sub in subList:
        if "READY" in sub:
            prefix = sub.replace("_READY", "")
            if appCounter <= xappssubmits:
                screenshot(prefix)
            else:
                return 0
        

"""
# SCREENSHOT AF PILOT - SÅDAN SOM DET SKAL SE UD
def run():
    i=0
    while (i < 10):
        for root, dirs, files in os.walk(r"C:\Users\David\Projekter\aMediaTest\Device\NytSetup\WindowsPhone\Customer"):
            for name in files:
                if name.endswith(".sln"):
                    mainflow(root,name)
                    i+=1
"""

def startCombination():
    combinePrefixWithIsSubmit()
    global subList
    subList = submissionList()

def resetMarkerInGoogleDocs():
    switchTo(googledocs)
    sleep(5)

    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)

def appIsReady():
    if any("READY" in sub for sub in subList):
        return True
    else:
        return False

def checkList():
    try:
        startCombination()
    except:
        log.printMessage("Error in combination: " + str(sys.exc_info()[1]))
        #log.printMessage("ERROR - Failed submitting for " + name)
        restart()

    if(appIsReady() == True):
        runList(subList)
        global startUpFlag
        startUpFlag = False

    resetMarkerInGoogleDocs()

global startUpFlag
startUpFlag = True

global appCounter
appCounter = 1

def run():
    i=0
    while(i<3):
        checkList()
        i+=1
    if appCounter-1 == xappssubmits:
        printMessage("Maximum specified app screenshot reached..." + str(xappssubmits))
    printMessage("All WP apps should have screenshot now...")

run()
close()
