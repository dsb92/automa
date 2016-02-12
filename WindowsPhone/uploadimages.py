#!/usr/bin/python
# -*- coding: utf-8 -*-

from automa.api import *
from win32api import GetSystemMetrics
from time import sleep
import os, sys, shutil
import searchgraphics as graphics
import log as log
from parameters import *

urldocs = "https://docs.google.com/spreadsheets/d/1WhbzQ2XvqOo_6R-twjdm7Z8FT1AGrq9QHe-7Ze3VwYk/edit#gid=494592811"

"Directory from where this script is located"
rootdir = os.path.dirname(os.path.realpath(__file__))
os.chdir(rootdir)

global debugCounter
debugCounter = 13

global endFlag
endFlag = False

dummyPoint = Point(x=5, y=500)

class graphicException(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)

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
    log.printMessage("Adding prefixes from google docs to list...")
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
    log.printMessage("Adding submit status from google docs to list...")

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
    log.printMessage("Combining prefix list with submit status list...")
    i=0
    comb = []
    while(i<len(prefixes)):
        comb.append(str(prefixes[i])+"_"+str(ifsubmits[i]))
        i+=1

    return comb

def searchInDocsForPrefix(prefix):
    log.printMessage("Searching for prefix: " + prefix + " with submit status to list...")
    switchTo(googledocs)
    press(CTRL+'h')
    write(prefix)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)

def getMetadataFromGoogleDocs(tabnum, metadata):
    log.printMessage("Getting " + metadata + " from Google docs...'")
    switchTo(googledocs)
    SETAUTO(False)
    typeNum(tabnum, TAB)
    press(ENTER)
    SETAUTO(True)
    press(CTRL + 'a'), press(CTRL + 'c')
    log.printMessage("Succesfully copied " + metadata)
    
    switch_to(tempfileName)

    press(CTRL+'v')
    write(";")
    press(ENTER, CTRL+'s')
            
    findTempLocation(tempfileName)
    
    datafile = file(tempLocation)
    for line in datafile:
        if ";" in line:
            line = line.replace(";", "")
            line = line.strip()
            log.printMessage("Succesfully found " + metadata + ": " + line)
            press(CTRL+'a',DEL, CTRL+'s')
            return line


def SETAUTO(x):
    Config.auto_wait_enabled = x


def switchTo(window):
    log.printMessage("Switching to: " + window)
    switch_to(window)
    
    press(LWIN+UP)

def typeNum(num, whattype):
    log.printMessage("Tabbing, pressing down og up " + str(num) + " times...")
    c = 0
    while (c < num):
        press(whattype)
        c += 1

def shiftTab(num):
    log.printMessage("Shift tabbing " + str(num) + " times...")
    c = 0
    while (c < num):
        press(SHIFT+TAB)
        c += 1

def goToDashboard():
    switchTo(dashboardWindow)
    log.printMessage("Clicking Dashboard")
    press(HOME)
    if Text("Dashboard").exists() == False:
        typeNum(2, TAB)
        press(ENTER)
    else:
        click("Dashboard")


def searchTabAppName():
    appNameLower = appName.lower
    thirdLetter = appNameLower()[2]

    if appName.decode('iso-8859-1') == "Firda" or appName.decode('iso-8859-1') == "Romerikes Blad":
        typeNum(16, TAB)
        press(ENTER)
    elif appNameLower()[0] >= 'f':
        typeNum(11, TAB)
        press(ENTER)
    else:
        typeNum(6, TAB)
        press(ENTER)

def clickAndFindApp():
    try:
        wait_until(Text("Submit App").exists, timeout_secs=15)
    except:
        log.printMessage("TimeoutExpiredException: Text 'App'")

    log.printMessage("Finding " + appName + " from Dashboard...")

    try:
        if Text("Apps").exists() == False:
            sleep(5)
            SETAUTO(False)
            typeNum(20, TAB)
            press(ENTER)
            SETAUTO(True)
        else:
            click("Apps")
    except:
        SETAUTO(False)
        typeNum(20, TAB)
        press(ENTER)
        SETAUTO(True)

    sleep(15)

    try:
        click(Text(appName.decode("iso-8859-1")))
        log.printMessage("Clicking App name")
    except:
        if not TextField("Find my apps").exists():
            log.printMessage("Find my apps does not exists..Tabbing to App name")
            SETAUTO(False)
            typeNum(29, TAB)
            SETAUTO(True)
            write(appName.decode("iso-8859-1"))
              
        else:
            log.printMessage("Writing App name into 'Find my apps' textfield")
            click("Find my apps")
            write(appName.decode("iso-8859-1"), into="Find my apps")
        sleep(10)
        press(ENTER)
        sleep(5)
        press(ENTER)
        sleep(10)
        
        searchTabAppName()
        sleep(10)


def goToUploadAndDescribe():
    global tabbedFlag
    tabbedFlag = False
    if Text("Lifecycle").exists() == False:
        sleep(5)
        SETAUTO(False)
        typeNum(26,TAB)
        press(ENTER)
        SETAUTO(True)
        global tabbedFlag
        tabbedFlag = True
    else:
        click(Text("Lifecycle"))

    if Text("Complete").exists() == False:
        sleep(5)
        if tabbedFlag == True:
            typeNum(7,TAB)
            press(ENTER)
        else:
            SETAUTO(False)
            typeNum(33,TAB)
            press(ENTER)
            SETAUTO(True)
    else:
        click(Text("Complete"))

    try:
        wait_until(Text("Required").exists, timeout_secs=30)
        requiredTextPtn = Text("Required")
        log.printMessage("Clicking 'Upload and describe your package'")
        btnAppUpload = Point(x=requiredTextPtn.center.x, y=requiredTextPtn.y+65+93)
        click(btnAppUpload)
        sleep(15)
        SETAUTO(False)
        typeNum(39,TAB)
        SETAUTO(True)
        press(PGDN)
    except:
        log.printMessage("TimeoutExpiredException: Text 'Required'")
        log.printMessage("Tabbing to 'Upload and describe your package'")
        sleep(5)
        SETAUTO(False)
        typeNum(21,TAB)
        press(ENTER)
        SETAUTO(True)
        sleep(15)
        SETAUTO(False)
        typeNum(39,TAB)
        SETAUTO(True)
        press(PGDN)

                
def addImages(path):
    try:
        if os.path.isdir(path):
            click(Button("Tidligere placeringer").center - (10,0))  
            write(path)
            press(ENTER)
            
            log.printMessage("Adding images to screenshotfolder...")
            click(graphics.splashScreenImage)
            press(CTRL+'a')
            press(CTRL+'a')
            sleep(3)
            press(ENTER)
        else:
            log.printMessage("Folder: " + os.path.basename(path) + " does not exists!")
            press(ENTER)
    except:
        log.printMessage("Failed in addImages function...trying to tab to button 'placering'")
        typeNum(4, TAB)
        press(ENTER)
        write(path)
        press(ENTER)
        log.printMessage("From exception: Adding images to screenshotfolder...")
        click(graphics.splashScreenImage)
        press(CTRL+'a')
        press(CTRL+'a')
        press(ENTER)
        

        
def copyFromGraphicPath(root, prefix):
    graphicPath = graphics.getGraphicPath(rootdir, appName)
    screenshotPath = graphics.getScreenshotPath(root)
    screenshotPath += prefix
    if graphicPath is not None:
        os.chdir(root)
        if not os.path.exists(screenshotPath):
            log.printMessage("Creating empty screenshotfolder to contain the images...")
            os.makedirs(screenshotPath)

        shutil.copy(graphicPath+"/"+graphics.productImage, screenshotPath)
        shutil.copy(graphicPath+"/"+graphics.splashScreenImage, screenshotPath)
        addImages(screenshotPath)
        log.printMessage("Adding screenshots...")
        sleep(30)

        return True
    else:
        return False


def uploadImages(root, prefix):
    os.chdir(rootdir)
    execfile('opendialog.py')
    write(root)
    press(ENTER)
    if copyFromGraphicPath(root, prefix) == False:
        log.printMessage("Graphic path not found...skipping this app submit")
        press(ALT+F4)
        click(dummyPoint)
        press(HOME)
        raise graphicException("Graphic path not found...")            
            
    switchTo(dashboardWindow)
    click(dummyPoint)
            
    log.printMessage("Successfully added all images")


def saveFromAppUploadAndDesc():
    press(PGDN)
    if Text("Save").exists() == False:
        SETAUTO(False)
        typeNum(37, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Save")

    log.printMessage("Saved package upload and description")
    sleep(25)


def reviewAndSubmit():
    try:
        wait_until(Text("Required").exists, timeout_secs=30)
        log.printMessage("Saved package upload and description")
        log.printMessage("Required texts exists...")
    except:
        log.printMessage("Required texts DOES NOT exists...")
        sleep(sleepTime/2)
        
    press(PGDN)
    if Text("Review and submit").exists() == False:
        log.printMessage("Tabbing to Review and submit")
        SETAUTO(False)
        typeNum(20, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        log.printMessage("Clicking Review and submit")
        click("Review and submit")
    sleep(sleepTime/2)


def finalAppSubmit():
    log.printMessage("Submitting app...")
    press(PGDN, PGDN)

    try:
        click("Submit")
        sleep(15)
    except:
        SETAUTO(False)
        typeNum(25, TAB)
        press(ENTER)
        sleep(15)
        SETAUTO(True)


def goToLifeCyclePage():
    SETAUTO(False)
    typeNum(19,TAB)
    press(ENTER)
    sleep(5)
    SETAUTO(True)


def goToProducts():
    log.printMessage("Going to products")
    global tabbedFlag
    tabbedFlag = False
    
    if Text("Products").exists():
        click(Text("Products"))
        log.printMessage("Clicking products")
        sleep(5)
    else:
        global tabbedFlag
        tabbedFlag = True
        SETAUTO(False)
        typeNum(31,TAB)
        press(ENTER)
        SETAUTO(True)
        sleep(5)


def searchInProducts():
    log.printMessage("Seaching in products...")
    try:
        click(Text(InAppAlias.decode("utf-8")))
    except:
        if tabbedFlag == True:
            typeNum(7,TAB)
            press(ENTER)
        else:
            SETAUTO(False)
            typeNum(38,TAB)
            press(ENTER)
            SETAUTO(True)

    log.printMessage("Going to complete...")
    if Text("Complete").exists() == False:
        sleep(5)
        SETAUTO(False)
        typeNum(30,TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Complete")


def goToInAppProperties():   
    try:
        wait_until(Text("Required").exists, timeout_secs = 30)
        required = Text("Required")
        global requiredTextPtn
        requiredTextPtn = Point(x=required.center.x, y=required.y+65)
        click(requiredTextPtn)
        sleep(10)
    except:
        SETAUTO(False)
        typeNum(19, TAB)
        press(ENTER)
        SETAUTO(True)
        sleep(10)


def choosePriceForInAppProduct():
    log.printMessage("Choosing price for In-app product...")
    SETAUTO(False)
    typeNum(27, TAB)
    press(ENTER)
    typeNum(9, DOWN)
    press(ENTER)
    SETAUTO(True)


def saveInAppProductProperties():
    press(PGDN)
    if Text("Save").exists() == False:
        press(PGDN)
    else:
        click("Save")

    if Text("Save").exists() == False:
        SETAUTO(False)
        typeNum(16+53, TAB) #16 til 'Tag' og 53 ned til 'Save'
        press(ENTER)
        SETAUTO(True)
    else:
        click("Save")

    log.printMessage("Saving In-app properties...")

           
def goToInAppDescription():
    log.printMessage("Going to In-app product description...")
    try:
        wait_until(Text("Required").exists, timeout_secs=30)
        required = Text("Required")
        requiredTextPtn = Point(x=required.center.x, y=required.y)
        btnDescription = Point(x=requiredTextPtn.x, y=requiredTextPtn.y+65+93)
        click(btnDescription)
        log.printMessage("Going to In-app description")
        sleep(10)
    except:
        log.printMessage("TimeoutExpiredException/NameError: Text 'Required'")
        log.printMessage("Tabbing to 'In App description'")
        SETAUTO(False)
        typeNum(20,TAB)
        press(ENTER)
        SETAUTO(True)
        log.printMessage("Going to In-app description")
        sleep(10)


def addProductImageDialog():
    open_dialog = Window('Åbn')
    if open_dialog.exists():
        switch_to(open_dialog)
        log.printMessage("Open dialog is open. Adding Product Image file...")
        return True
    else:
        return False


def getMonitorResolution():
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)
    return Point(x=width, y=height)


def chooseProductImage():
    scroll_down(2)
    try:
        wait_until(Text("Languages").exists, timeout_secs=10)
        log.printMessage("Clicking coord for product image")
        i=0
        while (addProductImageDialog() == False and i<3):
            click(Text("Languages").center+(-30,522))
            sleep(5)
            i+=1

        if Window('Åbn').exists() == False:
            raise Exception("ERROR: Dialog product image is not open...")

        write(graphics.productImage, into="Filnavn:")
        press(ENTER)
        log.printMessage("Inserted Product Image")
    except:
        log.printMessage("Getting HARDCODED Coord for Product Image dialog...")
        i=0
        while (addProductImageDialog() == False and i<3):
            if getMonitorResolution() == Point(x=1920, y=1080):
                click(Point(x=500, y=685))
                sleep(5)
            elif getMonitorResolution() == Point(x=1366, y=768):
                click(Point(x=225, y=680))
                sleep(5)
            i+=1

        if Window('Åbn').exists() == False:
            raise Exception("ERROR: Dialog product image HARDCODED is not open...")

        write(graphics.productImage, into="Filnavn:")
        sleep(3)
        press(ENTER)
        log.printMessage("Inserted Product Image")
        sleep(10)
        

def saveInAppDescription():
    if Text("Save").exists() == False:
        press(PGDN)
        try:
            click("Save")
            log.printMessage("Saving In-app description...")
        except:
            SETAUTO(False)
            typeNum(18,TAB)
            press(ENTER)
            SETAUTO(True)
    else:
        click("Save")
        log.printMessage("Saving In-app description...")


def submitInAppProduct():
    if Text("Submit").exists() == False:
        typeNum(24,TAB)
        press(ENTER)
        log.printMessage("Submitting In App Product")
    else:
        click("Submit")
        log.printMessage("Submitting In App Product")


def printResults():
    log.printMessage("DONE - Successfully submitted app and in-app product for: " + appName)


def restart():
    switchTo(googledocs)
    sleep(5)

    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)
    switch_to(tempfileName)
    press(CTRL+'a',DEL,CTRL+'s')
    global startUpFlag
    startUpFlag = False

    open_dialog = Window('Åbn')
    if open_dialog.exists():
        switch_to(open_dialog)
        press(ALT+F4)

    if Window('Søg').exists():
        switch_to('Søg')
        press(ALT+F4)

    """
    try:
        while Window("Google Chrome").exists():
            switch_to("Google Chrome")
            press(ALT+F4)
    except:
        pass

    start("Google Chrome", '-new-window', "https://dev.windowsphone.com/en-us/dashboard?logged_in=1")
    sleep(5)
    start("Google Chrome", '-new-window', urldocs)
    sleep(10)
    """


def resetGoogleDocs():
    switchTo(googledocs)
    shiftTab(3)
    press(ENTER)
    press(CTRL+'a')
    write("SUBMITTED")
    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11,TAB)
    press(ENTER)
    SETAUTO(True)


def close():
    if Window(tempfileName).exists():
        switch_to(tempfileName)
        press(ALT+F4)

    if Window(tempfileName2).exists():
        switch_to(tempfileName2)
        press(ALT+F4)

    if Window("Microsoft Visual Studio").exists():
        kill("Microsoft Visual Studio")
        sleep(3)

    SETAUTO(False)
    try:
        while Window("Google Chrome").exists():
            kill("Google Chrome")
    except:
        pass
    SETAUTO(True)

def appmainflow(root,name,prefix):
    try:
        log.printMessage("*****************APP SUBMIT NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appName
        appName = getMetadataFromGoogleDocs(4, "App name")
        goToDashboard()
        clickAndFindApp()
        goToUploadAndDescribe()
        uploadImages(root, prefix)
        saveFromAppUploadAndDesc()
        reviewAndSubmit()
        # OBS OBS
        #finalAppSubmit()
        #goToLifeCyclePage()
        goToDashboard()
        clickAndFindApp()
        goToProducts()
        searchInProducts()
        #goToInAppProperties()
        #choosePriceForInAppProduct()
        #saveInAppProductProperties()
        goToInAppDescription()
        chooseProductImage()
        saveInAppDescription()
        # OBS OBS
        #submitInAppProduct()
        resetGoogleDocs()
        printResults()
        
        global startUpFlag
        startUpFlag = False
        log.printMessage("*****************END OF APP SUBMIT NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appCounter
        appCounter += 1
    except graphicException as e:
        print 'FAILED - graphics: ', e.value
        restart()
        
    except:
        print "Unexpected error:", sys.exc_info()[1]
        log.printMessage("ERROR - Failed submitting for " + name)
        restart()
        

def foreach(prefix):
    for root, dirs, files in os.walk(rootdir+"/"+prefix):
        for name in files:
            if name.endswith(".sln"):
                appmainflow(root,name,prefix)
                return 0


def submit(prefix):
    for dir in os.listdir(rootdir):
        prefix = prefix.replace("_WP8", "")
        if prefix == dir:
            global visiolink_project_name
            visiolink_project_name = prefix + "_WP8"
            searchInDocsForPrefix(visiolink_project_name)
            foreach(prefix)
            return 0



def runList(slist):
    for sub in subList:
        if "READY" in sub:
            prefix = sub.replace("_READY", "")
            if appCounter <= xappssubmits:
                submit(prefix)
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

def appIsReady():
    if any("READY" in sub for sub in subList):
        return True
    else:
        return False

def checkList():
    startCombination()

    if(appIsReady() == True):
        runList(subList)
        global startUpFlag
        startUpFlag = False

    resetMarkerInGoogleDocs()

global appCounter
appCounter = 1

global startUpFlag
startUpFlag = True

#Tjekker hele listen 3 gange for at tjekke at alt er gået godt.
def run():
    start("Google Chrome", '-new-window', "https://dev.windowsphone.com/en-us/dashboard?logged_in=1")
    sleep(10)
    
    i=0
    while(i<3):
        checkList()
        i+=1
    if appCounter == xappssubmits:
        log.printMessage("Maximum specified app submit reached..." + str(xappssubmits))
    log.printMessage("All apps are submitted...")

run()
close()


