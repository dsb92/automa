#!/usr/bin/python
# -*- coding: utf-8 -*-

from automa.api import *
from win32api import GetSystemMetrics
from time import sleep
import os, sys, shutil
import xmlparser
import searchgraphics as graphics
import log as log
from parameters import *

#OBS: Kør kun Use case 1: Resefrver app navn, hvis ikke allerede kørt.

urldocs = "https://docs.google.com/spreadsheets/d/1WhbzQ2XvqOo_6R-twjdm7Z8FT1AGrq9QHe-7Ze3VwYk/edit#gid=494592811"

"Directory from where this script is located"
rootdir = os.path.dirname(os.path.realpath(__file__))
os.chdir(rootdir)

global debugCounter
debugCounter = 13

global endFlag
endFlag = False

dummyPoint = Point(x=5, y=500)

def debugWaitForNextSubmit():
    raw_input("------------> Continue next app reserve?(Press Enter)\n")

def debugAppName():
    return "_AlphaTest" + str(debugCounter)

def debugInAppAlias():
    return "_AlphaTest" + str(debugCounter)

def debugInAppId():
    return "_AlphaTest" + str(debugCounter)

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


#Kun vindue/hjemmeside skal maksimeres, når der skiftes.
def switchTo(window):
    log.printMessage("Switching to: " + window)
    switch_to(window)
    
    press(LWIN+UP)

#Tab til højre
def typeNum(num, whattype):
    log.printMessage("Tabbing, pressing down og up " + str(num) + " times...")
    c = 0
    while (c < num):
        press(whattype)
        c += 1

#Tab til venstre
def shiftTab(num):
    log.printMessage("Shift tabbing " + str(num) + " times...")
    c = 0
    while (c < num):
        press(SHIFT+TAB)
        c += 1

"**********USE CASE 1: Reserve app navn and choose category*************"
def goToDashboard():
    switchTo(dashboardWindow)
    log.printMessage("Clicking Dashboard")
    press(HOME)
    if Text("Dashboard").exists() == False:
        typeNum(2, TAB)
        press(ENTER)
    else:
        click("Dashboard")
      
def clickSubmitApp():
    try:
        if Text(dashboardWindow).exists():
            log.printMessage("Dashboard Window exists")
            log.printMessage("Clicking 'Submit App'")
            click(Text("Submit App", above="Apps"))
        else:
            switchTo(dashboardWindow)
            SETAUTO(False)
            typeNum(18, TAB)
            press(ENTER)
            SETAUTO(True)
    except:
        switchTo(dashboardWindow)
        SETAUTO(False)
        typeNum(18, TAB)
        press(ENTER)
        SETAUTO(True)


#"private" function
def userWantsBeta():
    #LOGIK her til at checke om Bruger har valgt beta.
    if submittype == 'b':
        return True
    else:
        return False

#"private" function
def spanMoreOptions():
    SETAUTO(False)
    typeNum(31, TAB)
    press(ENTER)
    SETAUTO(True)
    press(PGUP)

def goToAppInfo():
    try:
        wait_until(Text("Required").exists, timeout_secs=15)
        required = Text("Required")
        global requiredTextPtn
        requiredTextPtn = Point(x=required.center.x, y=required.y+65)
        btnAppInfo = requiredTextPtn
        log.printMessage("Clicking 'App Info'")
        click(btnAppInfo)
        sleep(10)
    except:
        sleep(5)
        SETAUTO(False)
        typeNum(20, TAB)
        log.printMessage("Tabbing to 'App Info'")
        press(ENTER)
        SETAUTO(True)
        sleep(10)

    """    
    if userWantsBeta():
        try:
            wait_until(Text("Name*").exists, timeout_secs=15)
        except:
            sleep(10)  

        spanMoreOptions()
    """
    
def ReserveAppName():
    switchTo(dashboardWindow)
    try:
        click(TextField("Name*")), press(CTRL+'v')
    except:
        SETAUTO(False)
        typeNum(20, TAB)
        SETAUTO(True)
        press(CTRL+'v')
        
    # OBS OBS
    #write(debugAppName())

    press(SHIFT+TAB)
    press(TAB)
    log.printMessage("Reserving app name")
    press(TAB, ENTER)
    sleep(10)

    press(CTRL+'a', CTRL+'c')
    switch_to(tempfileName)
    press(CTRL+'v', CTRL+'s')

    findTempLocation(tempfileName)

    datafile = file(tempLocation)
    if any("Pick another name" in line for line in datafile):
        log.printMessage("WARNING: " + appName + " is not available!!")
        switchTo(googledocs)
        shiftTab(docsresetnum)
        press(ENTER)
        press(CTRL+'a')
        write("UNAVAILABLE")
        raise Exception("WARNING: " + appName + " is not available!!")

    press(CTRL+'a',DEL, CTRL+'s')
    switchTo(dashboardWindow)
    click(dummyPoint)
    press(PGDN)


def chooseCategory():
    if ComboBox("Category*").exists() == False:
        SETAUTO(False)
        typeNum(19, TAB)
        log.printMessage("Tabbing to Category")
        press(ENTER)
        SETAUTO(True)
    else:
        log.printMessage("Clicking Category")
        click(ComboBox("Category*"))

    press(PGDN)
    log.printMessage("Choosing news and weather category...")

    try:
        while ComboBox("Category*").value != "news + weather":
            press(UP)
    except:
        typeNum(5,UP)

    """
    # HVIS BRUGER HAR VALGT BETA APPS.
    if userWantsBeta():
        press(TAB)
        sleep(3)
        typeNum(10, TAB)
        press(DOWN)
        press(TAB)
        i = 0
        while(i <= len(participants)-1):
            write(participants[i]+";")
            i += 1
    else:
        click(dummyPoint)
    """
    click(dummyPoint)
    press(PGDN)


# TRY CATCH HER HVIS WAIT_UNTIL SMIDER EXCEPTION - OK
def saveAppInfo():
    if Text("Save").exists() == False:
        SETAUTO(False)
        typeNum(29, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Save")

    log.printMessage("Saving App info...")
    try:
        wait_until(Text("Required").exists, timeout_secs=15)
    except:
        sleep(10)
        switchTo(dashboardWindow)
"**********END OF USE CASE 1: Reserve app navn and choose category**********"

def searchTabAppName():
    appNameLower = appName.lower
    thirdLetter = appNameLower()[2]

    if appName.decode('iso-8859-1') == "Firda" or appName.decode('iso-8859-1') == "Romerikes Blad":
        typeNum(16, TAB)
        press(ENTER)
    elif appName.decode('iso-8859-1') == "Ås Avis":
        typeNum(6, TAB)
        press(ENTER)
    elif appNameLower()[0] >= 'f':
        typeNum(11, TAB)
        press(ENTER)
    else:
        typeNum(6, TAB)
        press(ENTER)

"**********USE CASE 2: Find and save publisher and product id***********"
# TRY CATCH HER HVIS WAIT_UNTIL SMIDER EXCEPTION - OK
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
        # OBS OBS OBS
        #click(Text(appName.decode("iso-8859-1")+debugAppName()))
        click(Text(appName.decode("iso-8859-1")))
        log.printMessage("Clicking App name")
    except:
        # OBS OBS OBS
        if not TextField("Find my apps").exists():
            log.printMessage("Find my apps does not exists..Tabbing to App name")
            SETAUTO(False)
            typeNum(29, TAB)
            SETAUTO(True)
            #write(appName.decode("iso-8859-1")+debugAppName())
            write(appName.decode("iso-8859-1"))
              
        else:
            log.printMessage("Writing App name into 'Find my apps' textfield")
            click("Find my apps")
            #write(appName.decode("iso-8859-1")+debugAppName(), into="Find my apps")
            write(appName.decode("iso-8859-1"), into="Find my apps")
        sleep(10)
        press(ENTER)
        sleep(5)
        press(ENTER)
        sleep(10)

        searchTabAppName()
        sleep(10)

def goToAppDetails():
    log.printMessage("Clicking App details")

    if Text("Products").exists() == False and Text("Details").exists() == False:
        SETAUTO(False)
        typeNum(30, TAB)
        press(ENTER)
        sleep(5)
        SETAUTO(True)
    else:
        click(Text("Details", to_left_of="Products"))
        sleep(5)


def findPublisherId():
    log.printMessage("Finding publisher id...")
    press(CTRL+'a', CTRL+'c')
    switch_to(tempfileName)
    press(CTRL+'v', ENTER, CTRL+'s')

    findTempLocation(tempfileName)
    
    datafile = file(tempLocation)
    for line in datafile:
        if "CN" in line:
            global publisherId
            publisherId = line.replace("CN=", "")
            publisherId = publisherId.strip() # Fjerner alt mellemrum i start og slut af strengen.
            log.printMessage("Succesfully found publisherId: " + publisherId)
            break

    press(CTRL+'a',DEL, CTRL+'s')
    switchTo(dashboardWindow)
    click(dummyPoint)


def findProductId():
    log.printMessage("Finding product id...")
    press(CTRL+'a', CTRL+'c')
    switch_to(tempfileName)
    press(CTRL+'v', CTRL+'s')

    press(CTRL+'b')
    switch_to("Søg")
    write("App ID")
    try:
        click(RadioButton("Tilbage"))
    except:
        typeNum(2, TAB)
        press(LEFT)
        
    press(TAB,ENTER,TAB,ENTER,DEL,DEL)
    write(";"), press(CTRL+'s')

    findTempLocation(tempfileName)
    
    datafile = file(tempLocation)
    for line in datafile:
        if ";" in line:
            global productId
            productId = line.replace(";", "")
            productId = productId.strip()
            log.printMessage("Succesfully found product id: " + productId)
            break

    press(CTRL+'a', DEL, CTRL+'s')
    switchTo(dashboardWindow)
    click(dummyPoint)

"**********END OF USE CASE 2: Find and save publisher and product id***********"

"**********USE CASE 3: Opdatere manifest filer i app solution projekt***"
# TRY CATCH HER HVIS WAIT_UNTIL SMIDER EXCEPTION - OK
def goToUploadAndDescribe():
    global tabbedFlag
    tabbedFlag = False
    if Text("Lifecycle").exists() == False:
        sleep(5)
        SETAUTO(False)
        typeNum(26, TAB)
        press(ENTER)
        SETAUTO(True)
        global tabbedFlag
        tabbedFlag = True
    else:
        click(Text("Lifecycle"))

    if Text("Complete").exists() == False:
        sleep(10)
        if tabbedFlag == True:
            typeNum(7, TAB)
            press(ENTER)
        else:
            SETAUTO(False)
            typeNum(33, TAB)
            press(ENTER)
            SETAUTO(True)
    else:
        click(Text("Complete"))

    try:
        wait_until(Text("Required").exists, timeout_secs=40)
        requiredTextPtn = Text("Required")
        log.printMessage("Clicking 'Upload and describe your package'")
        btnAppUpload = Point(x=requiredTextPtn.center.x, y=requiredTextPtn.y+65+93)
        click(btnAppUpload)
    except:
        log.printMessage("TimeoutExpiredException: Text 'Required'")
        log.printMessage("Tabbing to 'Upload and describe your package'")
        sleep(5)
        SETAUTO(False)
        typeNum(21, TAB)
        press(ENTER)
        SETAUTO(True)
        

def editWMAppManifest(root):
    log.printMessage("Editing WMAppManifest.xml file...")
    #productId = "415e3737-82d6-4759-9bde-f6b693a64b03"
    #publisherId = "5860D604-48AB-492F-9714-6351AE57306E"
    xmlparser.xmlParse_WMAppManifest(root, productId, visiolink_project_name, appName, publisherId)


def resetToAppNameInGoogleDocs(num):
    switchTo(googledocs)
    SETAUTO(False)
    shiftTab(num)
    SETAUTO(True)


def editSettingsXML(root, prefix):
    log.printMessage("Editing Settings.xml file...")
    
    newTrackingAccount = getMetadataFromGoogleDocs(6, "Tracking account")
    newPrivacyPolicy = getMetadataFromGoogleDocs(7, "Privacy policy")
    #newAboutLocation = getMetadataFromGoogleDocs(9, "About location")
    newAboutLocation = "http://device.e-pages.dk/content/amedia/webview.php?view=info&customer="+prefix
    #newWebURL = getMetadataFromGoogleDocs(3, "External Web URL")
    newWebURL = "http://device.e-pages.dk/content/amedia/webview.php?view=home&customer="+prefix+"&platform=windows&device_type=phone"
    newComscoreCustom = getMetadataFromGoogleDocs(17, "Comscore customer C2")
    newComscoreDomain = getMetadataFromGoogleDocs(1, "Comscore domain/app name")
    newComscorePubSecret = getMetadataFromGoogleDocs(1, "Comscore publisher secret")
    newComscoreVirtual = getMetadataFromGoogleDocs(1, "Comscore virtual")
    #newHelppage = getMetadataFromGoogleDocs(1, "Help page")
    newHelppage = "http://device.e-pages.dk/content/amedia/webview.php?view=help&customer="+prefix
    resetToAppNameInGoogleDocs(11)

    xmlparser.xmlParse_SettingsXML(root,
                                   newTrackingAccount,
                                   newPrivacyPolicy,
                                   newAboutLocation,
                                   newWebURL,
                                   appName,
                                   prefix,
                                   newComscoreCustom,
                                   newComscoreDomain,
                                   newComscorePubSecret,
                                   newComscoreVirtual,
                                   newHelppage)


def editSettingsXAML(root):
    log.printMessage("Editing Settings.xaml file...")

    newColorCodeTheme = getMetadataFromGoogleDocs(1, "Color code theme")
    xmlparser.xmlParse_SettingsXAML(root, newColorCodeTheme, appName)
    resetToAppNameInGoogleDocs(23)
    

def copySplashScreenToProject(root):
    log.printMessage("Copying SplashScreenImage.jpg to to project...")
    splash = "SplashScreenImage.jpg"
    graphicPath = graphics.getGraphicPath(rootdir,appName)

    if graphicPath == None:
        raise Exception("ERROR: Graphicpath not found.")

    os.chdir(graphicPath)
    shutil.copy(graphics.splashScreenImage, root)
    os.chdir(root)
    
    if os.path.exists(splash):
        os.remove(splash)

    os.rename(graphics.splashScreenImage, splash)


def copyTilesToProject(root):
    log.printMessage("Copying tiles to project...")
    appIcon = "ApplicationIcon.png"
    smallTile = "FlipCycleTileSmall.png"
    medTile = "FlipCycleTileMedium.png"
    assetTilePath = root + "/Assets/Tiles"
    appIconPath = root + "/Assets"
    # OBS OBS - hurtige og dumme løsning - tilføjet d. 20/11-2014.
    forbiddenPath = root + "/Winnie/Assets/Tiles"
    graphicPath = graphics.getGraphicPath(rootdir,appName)
    
    if graphicPath == None:
        raise Exception("ERROR: Graphicpath not found.")

    os.chdir(graphicPath)
    source = os.listdir(graphicPath)
    for files in source:
        if files.endswith(".png") and not "99x99" in files:
            destination = assetTilePath
            shutil.copy(files,destination)

        if files.endswith(".png") and "99x99" in files:
            destination = appIconPath
            shutil.copy(files,destination)

        if files.endswith(".png") and "202x202" in files:
            destination = forbiddenPath
            shutil.copy(files,destination)

    os.chdir(appIconPath)
    
    if os.path.exists(appIcon):
        os.remove(appIcon)
        os.rename(graphics.appIcon, appIcon)
    else:
        os.rename(graphics.appIcon, appIcon)
    
    os.chdir(assetTilePath)

    if os.path.exists(smallTile):
        os.remove(smallTile)
        os.rename(graphics.smallTile, smallTile)
    else:
        os.rename(graphics.smallTile, smallTile)

    if os.path.exists(medTile):
        os.remove(medTile)
        os.rename(graphics.medTile, medTile)
    else:
        os.rename(graphics.medTile, medTile)


    # OBS OBS
    os.chdir(forbiddenPath)

    if os.path.exists(smallTile):
        os.remove(smallTile)
        os.rename(graphics.smallTile, smallTile)
    else:
        os.rename(graphics.smallTile, smallTile)


    os.chdir(graphicPath)
    source = os.listdir(graphicPath)
    for files in source:
        if files.endswith(".png"):
            destination = assetTilePath
            shutil.copy(files,destination)


def editCSPROJ(root):
    global xapversion
    xapversion = xmlparser.getVersionOfXAP(root)

    xmlparser.xmlParse_CSPROJ(root, xapversion)


#TRY CATCH HER HVIS WAIT_UNTIL SMIDER EXCEPTION - OK
def openslnfromVS2013(root, name):
    if startUpFlag == True:
        log.printMessage("Starting Visual Studio 2013")
        start("Visual Studio 2013")
    else:
        try:
            log.printMessage("Switching to Visual Studio 2013")
            switchTo("Microsoft Visual Studio")
        except:
            log.printMessage("Starting Visual Studio")
            start("Visual Studio 2013")

    log.printMessage("Opening solution project in VS2013...")
    try:
        click("FILE"), hover("Open"), click("Project/Solution...")
    except:
        sleep(sleepTime/2)
        press(CTRL+SHIFT+'o')

    if Button("Tidligere placeringer").exists() == False:
        typeNum(4, TAB)
        press(ENTER)
    else:
        click(Button("Tidligere placeringer").center - (10,0))

    write(root)
    press(ENTER)
    #doubleclick(name)
    #waint_until(Text("Ready").exists, timeout_secs=300, interval_secs=3.0)

    try:
        write(name, into="Filnavn:")
        sleep(3)
        press(ENTER)

    except:
        typeNum(7, TAB)
        press(ENTER)
        write(name)
        sleep(3)
        press(ENTER)
        
    sleep(sleepTime/2)


def buildNewReleaseAndClose():
    switchTo("Microsoft Visual Studio")
    sleep(5)
    click("Solution Configurations"), click("Release")
    sleep(5)
    click("Solution Platform")
    hover("ARM"), click("ARM")
    sleep(sleepTime)
    press(F6)
    sleep(25)
    press(F6)
    sleep(25)
    switchTo(dashboardWindow)

"**********END OF USE CASE 3: Opdatere manifest filer i app solution projekt***"   


"**********USE CASE 4: Upload og beskrivelse af XAP fil**********"

def getXAPFile(root):
    for filename in os.listdir(root+xapfolder):
        if filename.endswith(".xap"):
            version_current_date = "_"+str(xapversion)+"_"+str(xmlparser.current_date)
            if version_current_date in filename:
                xapfile = filename
                log.printMessage("Found XAP file: " + xapfile)
                return xapfile


def getMonitorResolution():
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)
    return Point(x=width, y=height)


def findAddNewBtn():
    save = "Save"
    upldescpack = "Upload and describe your package"
    if Text(save).exists():
        log.printMessage("Clickin  'Add new' relative to existing element: Save button")
        addnewPoint = Point(x=Text(save).center.x, y=Text(save).y-65)
        click(addnewPoint)
        sleep(5)
        
    elif Text(appName.decode('iso-8859-1')).exists():
        log.printMessage("Clicking 'Add new' relative to existing element: App name")
        addnewPoint = Point(x=Text(appName).left.x+25, y=Text(appName).left.y+304)
        click(addnewPoint)
        sleep(5)
        
    elif Text(upldescpack).exists():
        log.printMessage("Clicking 'Add new' relative to existing element: Upload and describe your package")
        addnewPoint = Point(x=Text(upldescpack).left.x+40, y=Text(upldescpack).left.y+345)
        click(addnewPoint)
        sleep(5)
        
    else:
        log.printMessage("!!!!!Clicking 'Add new' relative to existing element: HARDCODED COORD!!!!")
        if getMonitorResolution() == Point(x=1920, y=1080):
            click(Point(x=510, y=570))
            sleep(5)
        elif getMonitorResolution() == Point(x=1366, y=768):
            click(Point(x=240, y=570))
            sleep(5)
    

def addNewDialog():
    open_dialog = Window('Åbn')
    if open_dialog.exists():
        switch_to(open_dialog)
        log.printMessage("Open dialog is open. Adding XAP file...")
        return True
    else:
        return False


def spanMoreOptionsPerLang():
    if Text("More options per language").exists() == False:
        SETAUTO(False)
        typeNum(32, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("More options per language")


def getMonitorResolution():
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)
    return Point(x=width, y=height)


def addNewXAP(root):
    xapfile = getXAPFile(root)
    
    if xapfile == None:
        raise Exception("ERROR: No XAP file found in Release folder...")

    log.printMessage("Adding XAP file...")
    findAddNewBtn()

    i=0
    while(addNewDialog() == False and i<3):
        log.printMessage("Open dialog is NOT open. Retrying..."+str(i))
        findAddNewBtn()
        i+=1

    if addNewDialog() == False:
        raise Exception("ERROR: Cannot not add new XAP file because open dialog does not exists!!")

    try:
        if Button("Tidligere placeringer").exists() == False and Window('Åbn').exists():
            typeNum(4, TAB)
            press(ENTER)
        elif Button("Tidligere placeringer").exists() == True and Window('Åbn').exists():
            click(Button("Tidligere placeringer").center - (10,0))

        write(root+xapfolder)
        press(ENTER)
        write(xapfile, into="Filnavn:")
        sleep(3)
        press(ENTER)
    except:
        typeNum(4, TAB)
        press(ENTER)
        write(xapfile, into="Filnavn:")
        sleep(3)
        press(ENTER)

    sleep(sleepTime/2)
    log.printMessage("Successfully added XAP file")
    switchTo(dashboardWindow)
    click(dummyPoint)
    press(PGDN)

    spanMoreOptionsPerLang()
    
      
"""
def addNewXAP(root):
    xapfile = getXAPFile(root)

    if xapfile == None:
        raise Exception("ERROR: No XAP file found in Release folder...")
    
    try:
        log.printMessage("Adding XAP file...")
        try:
            addnewPoint = Point(x=Text("Save").center.x, y=Text("Save").y-65)
            click(addnewPoint)
        except:
            log.printMessage("LookUpException: Text 'Save'")
            if Text(appName).exists():
                text = "Upload and describe your package"
                if Text(text).exists():
                    click(Point(x=Text(text).left.x+40, y=Text(text).left.y+345))
                else:
                    log.printMessage("Clicking HARDCODED coord for 'Add new'")
                    click(Point(x=510, y=574))
            else:
                #OBS hardcodes....ingen andre løsninger. Skal helst ikke nå her.
                log.printMessage("Clicking HARDCODED coord for 'Add new'")
                click(Point(x=510, y=574))

        try:
            open_dialog = Window("Åbn")
            wait_until(lambda: open_dialog.exists() or Text("Filnavn:").exists(),
                       timeout_secs=2)
            
            if open_dialog.exists():
                switch_to(open_dialog)
 
            log.printMessage("Add new..")
        except:
            log.printMessage("Retrying...")
            
            open_dialog = Window("Åbn")

            if open_dialog.exists():
                switch_to(open_dialog)

        try:
            if Button("Tidligere placeringer").exists() == False and Window('Åbn').exists():
                typeNum(4, TAB)
                press(ENTER)
            else:
                click(Button("Tidligere placeringer").center - (10,0))
      
            for releaseFolder, dirs, files in os.walk(root+xapfolder):
                for name in files:
                    if name.endswith(".xap"):
                        write(releaseFolder)
                        press(ENTER)
                        write(xapfile, into="Filnavn:")
                        press(ENTER)
                        break
        except LookupError:
            log.printMessage("Package aldready added...")
            press(PGUP)
        
        sleep(sleepTime/2)
        log.printMessage("Successfully added XAP file")
        click(dummyPoint)
        press(PGDN)

        if Text("More options per language").exists() == False:
            SETAUTO(False)
            typeNum(31, TAB)
            press(ENTER)
            SETAUTO(True)
        else:
            click("More options per language")
        
    except:
        switchTo(dashboardWindow)
        if Window('Åbn').exists():
            switch_to(Window('Åbn'))
            press(ALT+F4)
        addNewXAP(root)
"""    

def getDescOfAppFromGoogleDocs():
    switchTo(googledocs)
    log.printMessage("Getting App Description from Google Docs...")
    SETAUTO(False)
    typeNum(9, TAB)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a',CTRL+'c')
    switchTo(dashboardWindow)
    
def insertDescFromApp():
    if Text("Description for the Store*").exists() == False:
        SETAUTO(False)
        typeNum(25, TAB)
        press(CTRL+'v')
        SETAUTO(True)
    else:
        write("", into="Description for the Store*")
        press(CTRL+'v')

    log.printMessage("Inserted App Description from Google Docs")

def getKeywordsForSearchFromGoogleDocs():
    log.printMessage("Getting search keywords from Google Docs")
    switchTo(googledocs)
    press(TAB)
    press(ENTER)
    press(CTRL+'a',CTRL+'c')
    switch_to(tempfileName)
    press(ENTER)
    press(CTRL+'v')
    write(",")
    press(UP)
    press(CTRL+'s')

    
def seperateKeywords():
    log.printMessage("seperating keywords...")
    datafile = open(tempLocation)
    global keywords
    keywords = []   

    for line in datafile:
        for letter in line:
            if ',' in letter:   
                press(CTRL+'b')
                switch_to("Søg")
                write(",")
                click(RadioButton("Fremad"))
                press(TAB,ENTER,TAB,ENTER)
                write(";")
                press(ENTER)

    press(CTRL+'s')

    datafile = open(tempLocation)
    
    for line in datafile:
        if ";" in line:
            keyword = line.replace(";", "")
            word = keyword.decode("iso-8859-1")
            word = word.strip()
            log.printMessage("Appending keyword: " + keyword)
            keywords.append(word)

    press(CTRL+'a',DEL,CTRL+'s')
    switchTo(dashboardWindow)


def insertKeywords():
    press(TAB,TAB)
    speckwText = "Specify keywords"
    global uploadImgPoint
    if Text(speckwText).exists() == False:
        #OBS OBS OBS
        uploadImgPoint = Point(x=515, y=492)
    else:
        uploadImgPoint = Point(x=Text(speckwText).center.x, y=Text(speckwText).y+325)

    SETAUTO(False)
    i = 0
    while(i <= len(keywords)-1): # len(keywords) - men hva så hvis keywords > 5?
        write(keywords[i])
        press(ESC) # Prevent website redirect.
        if i == maxAmountKeywords-1:
            break
        press(TAB)
        i += 1
    SETAUTO(True)

    if len(keywords) == 3:
        typeNum(1, TAB)

    elif len(keywords) == 2:
        typeNum(2, TAB)

    elif len(keywords) == 1:
        typeNum(3, TAB)
        

def getWebURLFromGoogleDocs():
    typeNum(2, TAB)
    weburl = getMetadataFromGoogleDocs(15, "Web URL for App")
    if not "http" in weburl:
        switchTo(dashboardWindow)
        write("http://")
        press(CTRL+'v')
    else:
        switchTo(dashboardWindow)
        press(CTRL+'v')

    log.printMessage("Added Web URL for App")

def getPrivacyURLFromGoogleDocs():
    log.printMessage("Getting Privacy URL from Google Docs...")
    typeNum(2, TAB)
    switchTo(googledocs)
    SETAUTO(False)
    shiftTab(12)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a',CTRL+'c')
    switchTo(dashboardWindow)
    press(CTRL+'v')
    log.printMessage("Added Privacy URL")
    
def getEmailFromGoogleDocs():
    log.printMessage("Getting E-Mail for support from Google Docs...")
    press(TAB)
    switchTo(googledocs)
    shiftTab(1)
    press(ENTER,CTRL+'a',CTRL+'c')
    switchTo(dashboardWindow)
    press(CTRL+'v')
    log.printMessage("Added E-Mail for support")
    press(PGDN)
    click(dummyPoint)

def saveFromAppUploadAndDesc():
    press(PGDN)
    if Text("Save").exists() == False:
        SETAUTO(False)
        typeNum(37, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Save")

    log.printMessage("Saving...")

# TRY CATCH HER HVIS WAIT_UNTIL SMIDER EXCEPTION
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
        
"**********END OF USE CASE 4: Upload og beskrivelse af XAP fil**********"

"""
"**********USE CASE 5: Tilføj In-app product**********"
def goToProducts():
    log.printMessage("Going to products")
    if Text("Products").exists():
        click(Text("Products"))
        log.printMessage("Clicking products")
        sleep(5)
    else:
        log.printMessage("Tabbing to products")
        typeNum(30, TAB)
        press(ENTER)
        sleep(5)

def addInAppProduct():
    if Text("Lifecycle").exists():
        click(Text("Lifecycle").center+(0,130))
        log.printMessage("Clicking Add in-app product..")
    else:
        log.printMessage("Tabbing to Add in-app product..")
        typeNum(31, TAB)
        press(ENTER)
        
def goToInAppProperties():
    log.printMessage("Going to In-app product properties...")
    try:
        wait_until(Text("Required").exists, timeout_secs=15)
        required = Text("Required")
        global requiredTextPtn
        requiredTextPtn = Point(x=required.center.x, y=required.y+65)
        click(requiredTextPtn)
        sleep(10)
    except:
        typeNum(19, TAB)
        press(ENTER)
        sleep(10)


                
def fillOutProductInfo():
    log.printMessage("Filling out In-app product properties...")
    #Unspan 'More options'
    press(PGDN)
    if Text("Save").exists() == False:
        typeNum(28, TAB)
        press(ENTER)
        press(HOME)
    else:
        click(Text("Save").center-(0,90))
        press(PGUP)
        
    
    typeNum(18, TAB)
    #In app alias
    #write(InAppAlias.decode("utf-8")+debugInAppAlias()) # OBS OBS Enkelt_kjøp_AppName.decode("iso-8859-1"), f.eks. Enkelt_kjøp_Haugesunds Avis
    write(InAppAlias.decode("utf-8")+"_"+appName.decode('iso-8859-1'))
    log.printMessage("Inserted In-app Alias")
    press(TAB)

    #In app product identifier
    #write(InAppProductIdentifier+debugInAppId()) # OBS OBS
    write(InAppProductIdentifier)
    log.printMessage("Inserted In-app product identifier")

    #Product type: Consumable default
    #press(TAB,TAB,ENTER,DOWN,ENTER)

    #Prduct lifetime if Durable
    #press(TAB,TAB,ENTER) + choose lifetime

    #Content type: Electronic software download - default
    #press(TAB,TAB,ENTER) + choose content type
    typeNum(5, TAB)
    press(ENTER)
    press(PGUP)
    typeNum(3, DOWN)
    press(ENTER)

    #Language - default norsk.
    typeNum(1, TAB)
    press(ENTER), press('n'), press(ENTER)
    log.printMessage("Chose NO language")
    #OBS OBS Choose pricing -  NOT AVAILABLE IN BETA??? 28 NOK - DEFAULT
    #press(TAB,TAB)

    #Tag
    if userWantsBeta():
        typeNum(15, TAB)
    else:
        typeNum(18, TAB)

    write(InAppTag)
    log.printMessage("Inserted In-app Tag")
    
def saveInAppProductProperties():
    if Text("Save").exists() == False:
        press(PGDN)
    else:
        click("Save")

    if Text("Save").exists() == False:
        press(PGDN)
    else:
        click("Save")
    
    if Text("Save").exists() == False:
        press(PGDN)
        typeNum(53, TAB)
        press(ENTER)

    log.printMessage("Saving In-app properties...")

# TRY CATCH HER HVIS WAIT_UNTIL SMIDER EXCEPTION - OK
def goToInAppDescription():
    try:
        wait_until(Text("Required").exists, timeout_secs=15)
        requiredTextPtn = Text("Required")
        btnDescription = Point(x=requiredTextPtn.center.x, y=requiredTextPtn.y+65+93)
        click(btnDescription)
        log.printMessage("Going to In-app description")
    except:
        log.printMessage("TimeoutExpiredException/NameError: Text 'Required'")
        log.printMessage("Tabbing to 'In-app description'")
        typeNum(20, TAB)
        press(ENTER)
        log.printMessage("Going to In-app description")
        
    
def chooseProductTitel():
    try:
        wait_until(Text("Languages").exists, timeout_secs=15)
        click(Text("Languages").center+(0,291))
    except:
        typeNum(19, TAB)

    #OBS OBS
    write(InAppAlias.decode("utf-8"))
    log.printMessage("Inserted InApp alias")
    press(TAB)

    
def saveInAppDescription():
    if Text("Save").exists() == False:
        press(PGDN)
        try:
            doubleclick("Save")
            log.printMessage("Saving In-app description...")
        except:
            typeNum(1, TAB)
            press(ENTER)
    else:
        doubleclick("Save")
        log.printMessage("Saving In-app description...")
"""    
        
def resetGoogleDocs():
    switchTo(googledocs)
    SETAUTO(False)
    shiftTab(docsresetnum)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a')
    write("READY")
    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)


def printResults():
    log.printMessage("DONE - Successfully uploaded app: " + appName)

"**********END OF USE CASE 5: Tilføj In-app product**********"

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

    if Window("Placeringen er ikke tilgængelig").exists():
        switch_to("Placeringen er ikke tilgængelig")
        press(ALT+F4)

    if Window("Gem som").exists():
        switch_to("Gem som")
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
    """
    sleep(10)
    
    
    # OBS OBS FJERN...
    global debugCounter
    debugCounter += 1

    """
    global appCounter
    appCounter += 1
    """
    


#OBS OBS de 3 parameter: root,name,prefix
def appmainflow(root,name,prefix):
    try:
        log.printMessage("*****************APP UPLOAD NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        #log.printMessage("*****************APP UPLOAD NR: " + str(appCounter) + "/" + str(xappssubmits)+"*****************")
        global appName
        appName = getMetadataFromGoogleDocs(4, "App name")
        goToDashboard()

        
        "**********USE CASE 1: Reserve app navn and choose category*************"
        clickSubmitApp()
        goToAppInfo()
        ReserveAppName()
        chooseCategory()
        saveAppInfo()
        
        "********************END OF USE CASE 1**********************************"

        "**********USE CASE 2: Find and save publisher, product id and app name***********"
        clickAndFindApp()
        goToAppDetails()
        findPublisherId()
        findProductId()
        
        "********************END OF USE CASE 2**********************************"
        
        
        "**********USE CASE 3: Opdatere manifest filer i app solution projekt***"
        goToUploadAndDescribe()
        copyTilesToProject(root)
        copySplashScreenToProject(root)
        editWMAppManifest(root)
        editSettingsXML(root, prefix)
        editSettingsXAML(root)
        editCSPROJ(root)
        openslnfromVS2013(root, name)
        buildNewReleaseAndClose()
        #debugWaitForNextSubmit()
        
        "********************END OF USE CASE 3**********************************"
        "**********USE CASE 4: Upload og beskrivelse af XAP fil*****************"
        addNewXAP(root)
        getDescOfAppFromGoogleDocs()
        insertDescFromApp()
        getKeywordsForSearchFromGoogleDocs()
        seperateKeywords()    
        insertKeywords()
        getWebURLFromGoogleDocs()
        getPrivacyURLFromGoogleDocs()
        getEmailFromGoogleDocs()
        saveFromAppUploadAndDesc()
        reviewAndSubmit()
        "********************END OF USE CASE 4**********************************"
            
        """
        "**********USE CASE 5: Tilføj In-app product****************************"
        goToDashboard()
        clickAndFindApp()
        goToProducts()
        addInAppProduct()    
        goToInAppProperties()
        fillOutProductInfo()
        saveInAppProductProperties()
        goToInAppDescription()
        chooseProductTitel()
        saveInAppDescription()
        "********************END OF USE CASE 5**********************************"
        """
        printResults()
        resetGoogleDocs()
        global startUpFlag
        startUpFlag = False
        sleep(5)
        log.printMessage("*****************END OF APP UPLOAD NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        #log.printMessage("*****************END OF APP RES NR: " + str(appCounter)+ "/" + str(xappssubmits)+"*****************")
        # OBS OBS ( FJERN... )
        
        global debugCounter
        debugCounter += 1
        global appCounter
        appCounter += 1
    except:
        log.printMessage("Unexpected error: " + str(sys.exc_info()[1]))
        #log.printMessage("ERROR - Failed submitting for " + name)
        restart()


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
    


def foreach(prefix):
    for root, dirs, files in os.walk(rootdir+"/"+prefix):
        for name in files:
            if name.endswith(".sln"):
                appmainflow(root,name,prefix)
                return 0



def submit(prefix):
    prefix = prefix.replace("_WP8", "")
    for dir in os.listdir(rootdir):
        if prefix == dir:
            global visiolink_project_name
            visiolink_project_name = prefix + "_WP8"
            searchInDocsForPrefix(visiolink_project_name)
            foreach(prefix)
            return 0


def runList(slist):
    for sub in subList:
        if "RESERVED" in sub:
            prefix = sub.replace("_RESERVED", "")
            if appCounter <= xappssubmits:
                submit(prefix)
            else:
                return 0


"""
#Kun til reservation af apps, husk ellers at udkommentere de tre overstående funktioner.     
def foreach(prefix):
    appmainflow()
    return 0
  
def runList(slist):
    for sub in subList:
        if "APP_NAME_RESERVED" in sub:
            if appCounter <= xappssubmits:
                prefix = sub.replace("_APP_NAME_RESERVED", "")
                searchInDocsForPrefix(prefix)
                foreach(prefix)
            else:
                return 0
"""

"""      
while(endFlag == False):
    combinePrefixWithIsSubmit()
    global subList
    subList = submissionList()
    if not any("RESERVED" in sub for sub in subList):
        global endFlag
        endFlag = True
        log.printMessage("All apps are ready to be submitted or already submitted...")
    else:
        if appCounter <= xappssubmits:
            runList(subList)
        else:
            global endFlag
            endFlag = True
            log.printMessage("Maximum specified app submit reached..." + str(xappssubmits))

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

def appIsReserved():
    if any("RESERVED" in sub for sub in subList):
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
        
    if(appIsReserved() == True):
        
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
    log.printMessage("All apps are ready to be submitted or already submitted...")

run()
close()
