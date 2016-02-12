#!/usr/bin/python
# -*- coding: utf-8 -*-

from automa.api import *
from time import sleep
import os, sys
from subprocess import call
import xmlparser
from random import randint
from datetime import datetime
from parameters import *

Config.search_timeout_secs = 60
Emulator = "Microsoft Windows Simulator"
folderOnDekstop = "ScreenshotWT"
current_date = datetime.now().date().isoformat()
wtsslog = "log_sreenshotWT_"+str(current_date)+".txt"

#Directory from where this script is located
rootdir = os.path.dirname(os.path.realpath(__file__))
sleepTime = 60

"****COORDINATION****"

#Config screenshot location
btnSettings = Point(x=921, y=327)
btnChooseSaveLocation = Point(x=1018, y=390)

#Capture screenshot
btnCapture = Point(x=920, y=300)

#Navigation
btnNextPage = Point(x=878, y=277)
btnBack = Point(x=72, y=87)
btnHome = Point(x=70, y=417)
btnHarPassOrd = Point(x=513, y=286)
btnSectionOverview = Point(x=851, y=123)

#Login
btnLogin = Point(x=550, y=482)
btnBackFromLogin = Point(x=551, y=61)
btnFieldLoginMobilNr = Point(x=613,y=172)
btnFieldLoginPass = Point(x=620, y=200)
# OBS OBS: HVIS VISIOLINK KODE ÆNDRER SIG, SKAL NEDSTÅENDE ÆNDRES!
txtFieldLoginMobilNr = "00009996"
txtFieldLoginPass = "test"
btnLoggInn = Point(x=767, y=170)
btnLoggOK = Point(x=656,y=296)

#ArticleView from SpreadView
btnRandomClick = Point(x=70,y=470)
btnZoomIn = Point(x=847, y=104)
btnZoomOut = Point(x=797, y=103)

def printMessage(msg, waitForUser=False):
    try:
        os.chdir(rootdir)
        f = open(wtsslog, 'a')
    
        if waitForUser:
            return raw_input ("------------> " + msg)
        else:
            now = datetime.now()
            nowtime = str(now.strftime("%H:%M:%S"))
            f.write(nowtime + "-------" + msg + "\n")
            print nowtime + "-------" + msg
            f.close()
    except:
        pass


def SETAUTO(x):
    Config.auto_wait_enabled = x

def typeNum(num, whattype):
    printMessage("Tabbing, pressing down or up " + str(num) + " times...")
    c = 0
    while (c < num):
        press(whattype)
        c += 1

#Kun vindue/hjemmeside skal maksimeres, når der skiftes.
def switchTo(window):
    switch_to(window)
    
    press(LWIN+UP)

def folderExists():
    if Window("Bekræft erstatning af mappe").exists():
        return True
    else:
        return False

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
            sleep(15)
            os.chdir(scripts_image_path)
            wait_until(Image("google_docs_W8.PNG").exists, timeout_secs=120)
            sleep(5)
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

def searchInDocsForPrefix(prefix):
    printMessage("Searching for prefix: " + prefix + " with submit status to list...")
    switchTo(googledocs)
    press(CTRL+'h')
    write(prefix)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)

def openSolutionFile(root, name):
    try:
        click("FILE"), hover("Open"), click("Project/Solution...")
    except:
        sleep(10)
        press(CTRL+SHIFT+'o')

    printMessage("Opening solution file...")
    if Button("Tidligere placeringer").exists() == False:
        typeNum(4, TAB)
        press(ENTER)
    else:
        click(Button("Tidligere placeringer").center - (10,0))

    write(root)
    press(ENTER)
    #doubleclick(name)
    write(name, into="Filnavn:")
    press(ENTER)

    sleep(sleepTime/2)

    """
    if Window("Microsoft Visual Studio").exists():
        kill("Microsoft Visual Studio")
        sleep(3)

    
    call(r"C:\Program Files (x86)\Common7\IDE\devenv.exe" + " " + root + "/" + name)
    sleep(sleepTime)
    """


def build(buildconfig):
    try:
        click("Solution Platforms")
    except:
        click(Point(x=572,y=60))

    if buildconfig == "ARM":
        typeNum(2, DOWN), press(ENTER)
        sleep(5)

    elif buildconfig == "x64":
        typeNum(5, DOWN), press(ENTER)
        sleep(5)

    elif buildconfig == "x86":
        typeNum(6, DOWN), press(ENTER)
        sleep(5)

    press(F6)
    sleep(20)
    #wait_until(Text("Build succeeded").exists, timeout_secs=30, interval_secs=3.0)
    printMessage("Build succeeded for " + buildconfig)
        

def buildSolution():
    printMessage("Building solution for all platforms...")
    try:
        click("Solution Configuration")
    except:
        click(Point(x=458,y=64))

    typeNum(2, DOWN), press(ENTER)
    sleep(5)

    build("ARM")
    build("x64")
    build("x86")
    

def setupEmulator():
    printMessage("Setting up Emulator...")
    try:
        click("Solution Platforms")
    except:
        click(Point(x=572,y=60))

    typeNum(3, DOWN), press(ENTER)

    try:
        click("Solution Configuration")
    except:
        click(Point(x=458,y=64))

    #Kør i Release
    typeNum(2, DOWN), press(ENTER)

    try:
        click(Text("Browser Link").center-(23,0))
        typeNum(2, DOWN), press(ENTER)
    except:
        printMessage("Could not choose emulator...guessing that 'Simulator' is default..")
        

def startWithoutDebugging():
    printMessage("Starting without debugging...")
    press(CTRL+F5)
    sleep(sleepTime/2)
    try:
        if Window("Lenovo Settings Power").exists():
            switch_to("Lenovo Settings Power")
            press(ALT+F4)

        printMessage("Switching to Emulator")
        switch_to(Emulator)
    except:
        switchTo("Microsoft Visual Studio")
        startWithoutDebugging()


def startUpConfigEmulator():
    printMessage("Dragging Emulator to upper left corner...")
    drag(Window(Emulator).left, Point(x=0, y=Window(Emulator).left.y))
    drag(Window(Emulator).top, Point(x=Window(Emulator).top.x, y=0))


def setResolution():
    btnResolution = Point(x=Window(Emulator).x+Window(Emulator).width-18, y=Window(Emulator).y+240)
    btnRes1366x768Type = Point(x=btnResolution.x+112, y=btnResolution.y+77)

    click(btnResolution)
    sleep(5)
    click(btnRes1366x768Type)

    drag(Window(Emulator).right, Point(x=2560, y=Window(Emulator).right.y)) 
    drag(Window(Emulator).left, Point(x=0, y=Window(Emulator).left.y))
    drag(Window(Emulator).top, Point(x=Window(Emulator).top.x, y=0))


def startUpSaveLocation():
    printMessage("Setting up save location...")
    click(btnSettings)
    try:
        click("Choose save location...")
    except:
        click(btnChooseSaveLocation)

    switch_to("Angiv en mappe")
    
    press(HOME)
    #os.chdir(scripts_image_path)
    #click(Image("desktop_icon.PNG"))
    click("Opret en ny mappe")
    write(folderOnDekstop)
    press(ENTER)

    if folderExists() == True:
        press(ENTER)

def newSaveFolder(ssfolder):
    if startUpFlag == True:
        click(folderOnDekstop)
    else:
        click(btnSettings)
        try:
            click("Choose save location...")
        except:
            click(btnChooseSaveLocation)
            
        switch_to("Angiv en mappe")
        click(folderOnDekstop)
          
    click("Opret en ny mappe")
    write(ssfolder), press(ENTER)
    if folderExists() == True:
        press(ENTER)
        press(DEL)
        click("Opret en ny mappe"), write(folderOnDekstop), press(ENTER)
        click("OK")
    else:
        click("OK")
        switch_to(Emulator)

def IfPurchasepage():
    os.chdir(scripts_image_path)
    if Image("purchase_page.PNG").exists():
        click(btnHarPassOrd)
        sleep(3)
        click(btnFieldLoginMobilNr)
        press(CTRL+'a'), write(txtFieldLoginMobilNr)
        press(TAB)
        press(CTRL+'a'), write(txtFieldLoginPass)
        click(btnLoggInn)
        sleep(6)


def beginCapturing(prefix):
    click(btnRandomClick)

    #choose todays paper
    press(TAB)
    press(ENTER)
    sleep(8)
    IfPurchasepage()

    #frontpage of paper - Screenshot 1
    printMessage("Capturing todays frontpage...")
    click(btnCapture)
    click(btnNextPage)
    sleep(8)

    #spreadview - Screenshot 2
    printMessage("Capturing spreadview...")
    click(btnCapture)

    #articleview - Screenshot 3
    xml = xmlparser.xmlParse_server_xml_todayspaper(prefix)

    #Hvis avis har sections og article view
    if xmlparser.xmlParse_server_has_sections(xml) and xmlparser.xmlParse_server_has_articles(xml):
        printMessage(prefix + " has sections and articles")
        btnArticleView = Point(x=865, y=470)

    #Hvis avis kun har article view (og pages)
    elif xmlparser.xmlParse_server_has_articles(xml):
        printMessage(prefix + " has articles only")
        btnArticleView = Point(x=865, y=415)

    #Hvis avis slet ikke har article view, tag ikke screenshot
    else:
        printMessage(prefix + " has NOT sections or articles")
        btnArticleView = None

    if btnArticleView is not None:
        rightclick(btnNextPage)
        sleep(8)
        click(btnArticleView)
        sleep(8)
        printMessage("Capturing articleview...")
        click(btnCapture)
    
    click(btnBack)
    sleep(8)
    
    #pageoverview - Screenshot 4
    # OBS OBS
    #click(btnCapture) # UDKOMMENTERE HER
    rightclick(btnBack)
    printMessage("Capturing pageview...")
    click(btnCapture)
    
    #startpage - Screenshot 5
    printMessage("Capturing kioskpage...")
    click(btnHome)
    sleep(8)
    click(btnCapture)


def resetGoogleDocs():
    printMessage("Resetting google docs...")
    switchTo(googledocs)
    press(TAB)
    press(ENTER)
    press(CTRL+'a')
    write("READY")
    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)

def restart():
    printMessage("RESTARTING......")
    os.chdir(rootdir)
    execfile('screenshot_log.py')
    
    global startUpFlag
    startUpFlag = False
    
    if Window("Angiv en mappe").exists():
        switch_to("Angiv en mappe")
        press(ALT+F4)

    if Window("Microsoft Visual Studio").exists():
        switchTo("Microsoft Visual Studio")

    if Window("Lenovo Settings Power").exists():
        switch_to("Lenovo Settings Power")
        press(ALT+F4)


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

    if Window("Lenovo Settings Power").exists():
        switch_to("Lenovo Settings Power")
        press(ALT+F4)


def mainflow(root,name,prefix):
    try:
        printMessage("*****************APP SCREENSHOT NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        
        if startUpFlag == True:
            VS2013 = start("Visual Studio 2013")
        

        openSolutionFile(root, name)
        buildSolution()
        setupEmulator()
        startWithoutDebugging()
                
        if startUpFlag == True:
            startUpConfigEmulator()
            #setResolution()
            startUpSaveLocation()

        os.chdir(root)
        os.chdir("..")
        mycwd = os.getcwd()
        screenshotFolder = os.path.basename(mycwd)
        newSaveFolder(screenshotFolder)
        beginCapturing(prefix)
        
        global startUpFlag
        startUpFlag = False
        printMessage("DONE - Finish capturing for: " + name)
        printMessage("*****************END OF APP SCREENHOT NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appCounter
        appCounter += 1
        searchInDocsForPrefix(visiolink_project_name)
        sleep(5)
        resetGoogleDocs()
    except:
        printMessage("Unexpected error: " + str(sys.exc_info()[1]))
        printMessage("ERROR - Failed screenshot for " + name)
        restart()


def ifPrefixExistsInPath(root):
    os.chdir(root)
    os.chdir("..")
    prefixpath = os.getcwd()
    prefixname = os.path.basename(prefixpath)
    userhome = os.path.expanduser('~')
    desktop = userhome + '/Desktop/'
    try:
        os.chdir(desktop+folderOnDekstop)
        if not any(prefixname in s for s in os.listdir(os.getcwd())):
            return False
        else:
            return True
    except:
        return False
    

def foreach(prefix):
    for root, dirs, files in os.walk(rootdir+"/"+prefix):  
        for name in files:
            if name.endswith(".sln") and ifPrefixExistsInPath(root) == False:
                mainflow(root,name,prefix)
                #print root
                return 0
            elif name.endswith(".sln") and ifPrefixExistsInPath(root) == True:
                return 0


def screenshot(prefix):
    for dir in os.listdir(rootdir):
        prefix = prefix.replace("_W8", "")
        if prefix == dir:
            global visiolink_project_name
            visiolink_project_name = prefix + "_W8"
            foreach(prefix)
            return 0


def runList(slist):
    for sub in subList:
        if "BUILT" in sub:
            prefix = sub.replace("_BUILT", "")
            if appCounter <= xappssubmits:
                screenshot(prefix)
            else:
                return 0


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


def appIsSubmitted():
    if any("BUILT" in sub for sub in subList):
        return True
    else:
        return False


def checkList():
    try:
        startCombination()
    except:
        printMessage("Error in combination: " + str(sys.exc_info()[1]))
        #log.printMessage("ERROR - Failed submitting for " + name)
        restart()

    if(appIsSubmitted() == True):
        runList(subList)
        global startUpFlag
        startUpFlag = False

    resetMarkerInGoogleDocs()
            

global startUpFlag
startUpFlag = True

global appCounter
appCounter = 1

def run():
    checkList()
    
    if appCounter-1 == xappssubmits:
        printMessage("Maximum specified app screenshot reached..." + str(xappssubmits))
    printMessage("All W8 apps should have screenshot now...")

run()
close()

