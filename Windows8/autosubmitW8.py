#!/usr/bin/python
# -*- coding: utf-8 -*-

from automa.api import *
from win32api import GetSystemMetrics
from subprocess import call
from random import randint
from time import sleep
import os, sys, shutil
from os.path import expanduser
import log as log
import xmlparser
import searchgraphics as graphics
from parameters import *

urldocs = "https://docs.google.com/spreadsheets/d/1e0sswHJo91jPbLfAdzTrC2dfFDbpGP6m8f81aMtP_2A/edit#gid=614091273"

"Directory from where this script is located"
rootdir = os.path.dirname(os.path.realpath(__file__))
os.chdir(rootdir)

menulistWindow = "Haugesunds Avis"
global debugCounter
debugCounter = 13

global endFlag
endFlag = False

dummyPoint = Point(x=5, y=500)

def debugWaitForNextSubmit():
    raw_input("------------> Continue next app submit?(Press Enter)\n")

def debugAppName():
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
    log.printMessage("Switching to " + window)
    switch_to(window)
    
    press(LWIN+UP)

#Tab til højre
def typeNum(num, whattype):
    log.printMessage("Tabbing or pressing " + str(num) + " times...")
    c = 0
    while (c < num):
        press(whattype)
        c += 1

#Tab til venstre
def shiftTab(num):
    log.printMessage("Shiftabbing or pressing " + str(num) + " times...")
    c = 0
    while (c < num):
        press(SHIFT+TAB)
        c += 1

def checkWindow():
    if Window(dashboardWindow).exists():
        switchTo(dashboardWindow)
    elif Window(sendEnAppWindow).exists():
        switchTo(sendEnAppWindow)
    elif Window(menulistWindow).exists():
        switchTo(menulistWindow)
    else:
        SETAUTO(False)
        try:
            while Window("Google Chrome").exists():
                kill("Google Chrome")
        except:
            pass
        SETAUTO(True)
        sleep(5)
        start("Google Chrome", '-new-window', "https://appdev.microsoft.com/StorePortals/da-DK/Home/Index")
        sleep(15)
        start("Google Chrome", '-new-window', urldocs)
        sleep(15)
        switchTo(dashboardWindow)
        

def checkImageExists():
    log.printMessage("Checking if image exists...")
    if Window("Gem som").exists():
        log.printMessage("Replacing image...")
        if not Button("Ja").exists():
            press(ENTER)
        else:
            click("Ja")

    if Window("Erstat filer eller spring dem over").exists():
        log.printMessage("Replacing image...")
        press(TAB,ENTER)


def IsMSSecurityWindow():
    #Hvis Microsoft login side dukker uforventet op.
    os.chdir(scripts_image_path)
    if Window("Log på din Microsoft-konto").exists() or Image("microsoft_security_window.PNG").exists():
        log.printMessage("WARNING: Microsoft security window is open...")
        write("MedApp2012")
        press(ENTER)
        try:
            wait_until(lambda: Window(dashboardWindow).exists() or
                       Image("sendEnAppWindow.PNG").exists() or
                       Window(sendEnAppWindow).exists() , timeout_secs=120)
            sleep(5)
            checkWindow()
            return True
        except:
           checkWindow()
           return True
    else:
        return False
        

def screenshotfolderExists():
    log.printMessage("Checking if screenshotfolder exists...")
    if Window("Bekræft erstatning af mappe").exists():
        log.printMessage("Replacing existing screenshotfolder...")
        click("Ja")
    else:
        log.printMessage("Screenshotfolder does not exists")
        press(ENTER)
            

"**********USE CASE 1: Reserve app navn**********"
def goToDashboard():
    if IsMSSecurityWindow() == False:
        checkWindow()
                
    log.printMessage("Clicking Dashboard")
    press(HOME)
    if Text("DASHBOARD").exists() == False:
        typeNum(6, TAB)
        press(ENTER)
    else:
        click("DASHBOARD")

    sleep(10)


def goToSendEnApp():
    log.printMessage("Clicking Send en app")

    if Text("Send en app").exists() == False:
        SETAUTO(False)
        typeNum(15, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Send en app")

def goToMenu_Appnavn():
    try:
        wait_until(Text("Appnavn").exists, timeout_secs=120)
    except:
        IsMSSecurityWindow()
        
    log.printMessage("Going to Menu: Appnavn")

    if Text("Appnavn").exists() == False:
        SETAUTO(False)
        typeNum(14, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Appnavn")


def Menu_Appnavn():
    global menulistWindow
    menulistWindow = appName.decode('iso-8859-1').encode('utf-8')
    
    try:
        os.chdir(scripts_image_path)
        wait_until(Image("sendEnAppWindow.PNG").exists, timeout_secs=120)
        switchTo(sendEnAppWindow)
    except:
        IsMSSecurityWindow()
        
    try:
        click(RadioButton("Brug et eksisterende appnavn"))
        press(TAB)
        press(ENTER)
        log.printMessage("Using existing appname...")
    except:
        SETAUTO(False)
        typeNum(17, TAB)
        SETAUTO(True)
        press(SPACE)
        press(TAB)
        press(ENTER)

    appNameLower = appName.lower
    firstLetter = appNameLower()[0]

    letterList = []

    os.chdir(rootdir)
    appnameList = file("appnames.txt")
    
    for appname in appnameList:
        appnameLetter = appname.lower()[0]
        if appnameLetter == firstLetter:
            log.printMessage("Appending letter..."+appnameLetter)
            letterList.append(appnameLetter)
        
    press(firstLetter.decode('iso-8859-1'))

    log.printMessage("Finding existing appname...")
    i = 1
    while ComboBox().value.encode('iso-8859-1').decode('iso-8859-1') != appName.decode('iso-8859-1'):
        if i > len(letterList):
            raise Exception("Existing appname not found...skipping to next app submit..")
        press(firstLetter.decode('iso-8859-1'))
        i += 1

    press(ENTER)
    press(TAB)
    press(ENTER)

    sleep(10)
    try:
        click("Gem")
        sleep(10)
    except:
        SETAUTO(False)
        typeNum(26, TAB)
        press(ENTER)
        SETAUTO(True)
        sleep(10)

"**********END OF USE CASE 1: Reserve app navn**********"        
    
"**********USE CASE 2: Salgsdetaljer**********"
def Menu_SalgsDetal():
    try:
        wait_until(Text("Salgsdetaljer").exists, timeout_secs=10)
    except:
        sleep(10)
        
    if Text("Salgsdetaljer").exists() == False:
        SETAUTO(False)
        typeNum(15, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Salgsdetaljer")

    log.printMessage("Going to Menu: Salgsdetaljer")
    
    sleep(10)
    
    #typeNum(26, TAB)
    press(ENTER)
    press(DOWN)
    press(ENTER)

    try:
        click("Markér alt")
    except:
        typeNum(6, TAB)
        
    press(ENTER)

    SETAUTO(False)
    typeNum(87, TAB)
    SETAUTO(True)

    press(ENTER)
    
    typeNum(8, DOWN)
    press(ENTER)
    sleep(3)
    typeNum(2, TAB)
    press(ENTER)
    press(DOWN)
    press(ENTER)
    sleep(3)

    typeNum(7, TAB)
    press(ENTER)
    
    sleep(10)

"**********END OF USE CASE 2: Salgsdetaljer**********"

"**********USE CASE 3: Tjenester**********"
   
def Menu_Tjenester():
    try:
        wait_until(Text("Tjenester").exists, timeout_secs=10)
    except:
        sleep(10)
        
    if Text("Tjenester").exists() == False:
        SETAUTO(False)
        typeNum(16, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Tjenester")


    log.printMessage("Going to Menu: Tjenester")
    
    sleep(10)

    SETAUTO(False)
    typeNum(31, TAB)
    SETAUTO(True)
    i = 0
    ones = 1
    tens = 0

    os.chdir(scripts_image_path)
    
    while(i<31):
        SETAUTO(False)
        if ones > 9:
            ones = 0
            tens += 1

        log.printMessage("Adding product " + InAppProductIdentifier+str(tens)+str(ones))
        write(InAppProductIdentifier+str(tens)+str(ones))
        
        if Image("chrome_crash.PNG").exists():
            restart()
            restartApplications()
        
        press(TAB,ENTER)
        typeNum(10, DOWN)
        press(ENTER)
        press(TAB,ENTER,PGUP)
        typeNum(2, DOWN)
        press(ENTER)
        
        press(TAB)
        press(ENTER)
        press(PGDN)
        press(ENTER)

        if i == 30:
            break

        press(TAB)
        sleep(3)
        press(ENTER)
            
        try:
            wait_until(Image("inn_app_product.PNG").exists, timeout_secs=20)
            sleep(5)
        except:
            sleep(10)

        sleep(5)
        ones += 1
        i += 1
        SETAUTO(True)
        
    typeNum(2, TAB)
    press(ENTER)
    sleep(10)
    
"**********END OF USE CASE 3: Tjenester**********"


"**********USE CASE 4: Alder-og klassifikationscertifikater**********"
def Menu_AlderOgKlass():
    try:
        wait_until(Text("Aldersklassifikation og klassifikationscertifikater").exists, timeout_secs=10)
    except:
        sleep(10)
        
    if Text("Aldersklassifikation og klassifikationscertifikater").exists() == False:
        SETAUTO(False)
        typeNum(17, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Aldersklassifikation og klassifikationscertifikater")


    log.printMessage("Going to Menu: Aldersklassifikation etc")
    sleep(10)
    
    SETAUTO(False)
    typeNum(2, DOWN)
    typeNum(16, TAB)
    press(ENTER)
    SETAUTO(True)
    sleep(10)
    
"**********END OF USE CASE 4: Alder-og klassifikationscertifikater**********"


"**********USE CASE 5: Kryptografi**********"
def Menu_Krypto():
    try:
        wait_until(Text("Kryptografi").exists, timeout_secs=10)
    except:
        sleep(10)
        
    if Text("Kryptografi").exists() == False:
        SETAUTO(False)
        typeNum(18, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Kryptografi")

    log.printMessage("Going to Menu: Kryptografi")
    
    sleep(10)

    press(DOWN)
    sleep(3)
    press(TAB)
    press(SPACE)
    typeNum(2, TAB)
    press(ENTER)
    sleep(10)

"**********END OF USE CASE 5: Kryptografi**********"


"**********USE CASE 6: Pakker**********"    
def resetToAppNameInGoogleDocs(num):
    switchTo(googledocs)
    SETAUTO(False)
    shiftTab(num)
    SETAUTO(True)


def getAPPXFile(path, buildconfig):
    for filename in os.listdir(path):
        if filename.endswith(".appx"):
            version_current_date = buildconfig+"_"+str(appxversion)+"_"+str(xmlparser.current_date)
            if version_current_date in filename:
                appxfile = filename
                log.printMessage("Found APPX file "+buildconfig+": " + appxfile)
                return appxfile
 

def openDialog():
    try:
        log.printMessage("Going to open dialog...")
        open_dialog = Window("Åbn")
        wait_until(lambda: open_dialog.exists() or Text("Filnavn:").exists(),
                   timeout_secs=5)

        if open_dialog.exists():
            switch_to(open_dialog)

    except:
        log.printMessage("Retrying to open dialog...")
        """
        SETAUTO(False)
        typeNum(26, TAB)
        press(ENTER)
        SETAUTO(True)
        """
        open_dialog = Window("Åbn")
        
        if open_dialog.exists():
            switch_to(open_dialog)
        else:
            raise Exception("ERROR: Open dialog is not open...")

        
    if Button("Tidligere placeringer").exists() == False and Window('Åbn').exists():
        typeNum(4, TAB)
        press(ENTER)
        log.printMessage("Tabbing to address line in open dialog...")
        
    elif Button("Tidligere placeringer").exists() == True and Window('Åbn').exists():
        log.printMessage("Clicking to address line in open dialog...")
        click(Button("Tidligere placeringer").center - (10,0))


def getPackageName(prefix, buildconfig, version):
    return prefix+"_"+buildconfig+"_"+version+"_"+str(xmlparser.current_date)


def removeIfAPPXExists(root, buildconfig):
    os.chdir(packagePathStore)
    source = os.listdir(os.getcwd())
    
    for files in source:
        if files.endswith(".appx"):
            version_current_date = buildconfig+"_"+str(appxversion)+"_"+str(xmlparser.current_date)
            if version_current_date in files:
                os.remove(files)
                break

def ClearPackageMap(root):
    pathToDel = root+metadataPath+"/package_maps"
    if os.path.exists(pathToDel):
        shutil.rmtree(pathToDel)
    

def buildAppxPackage(root, prefix, buildconfig):
    global appxversion
    appxversion = xmlparser.getVersionAppx(root)

    global packagename
    packagename = getPackageName(prefix, buildconfig, appxversion)

    global packagePathStore
    packagePathStore = root+metadataPath+"/AppxPackagesStore"

    if os.path.exists(packagePathStore) == False:
        os.mkdir(packagePathStore)
    else:
        removeIfAPPXExists(root, buildconfig)

    package_map_path_app = root+metadataPath+"/package_maps"
    package_map_path_root = rootdir + "/package_maps"+"/"+buildconfig+"/package.map.txt"

    if os.path.exists(package_map_path_app) == False:
        os.mkdir(package_map_path_app)
        os.chdir(package_map_path_app)
        os.mkdir(buildconfig)

        shutil.copy(package_map_path_root, os.getcwd()+"/"+buildconfig)
    else:
        if os.path.exists(package_map_path_app+"/"+buildconfig) == False:
            os.chdir(package_map_path_app)
            os.mkdir(buildconfig)
            
            shutil.copy(package_map_path_root, os.getcwd()+"/"+buildconfig)
            
    package_map_path_app = package_map_path_app + "/" + buildconfig
        
    os.chdir(package_map_path_app)

    with open("package.map.txt", 'r') as file:
        data = file.readlines()

    home = expanduser("~")

    i=0
    for d in data:
        data[i] = data[i].replace("PACKAGE_MAP_REPLACEMENT_PATH_HERE", rootdir)
        data[i] = data[i].replace("PREFIX_HERE", prefix)
        data[i] = data[i].replace("HOME_USER_PATH_HERE", home)
        i+=1

    with open("package.map.txt", 'w') as file:
        file.writelines(data)
    
    makeappxExePath = r"C:\Program Files (x86)\Windows Kits\8.1\bin\x64\makeappx.exe"
    buildmsg = "Building appx file for..."+buildconfig
    log.printMessage(buildmsg)
    call("echo "+buildmsg, shell=True)

    call(makeappxExePath + " " + "pack /l /h SHA256 /f " + package_map_path_app + "\package.map.txt /o /p " +
         packagePathStore + "/" + packagename + ".appx")


def signAppxPackage(root):
    signExe = r"C:\Program Files (x86)\Windows Kits\8.1\bin\x64\signtool.exe"

    call(signExe + " sign /fd SHA256 /a /f "+root+metadataPath+
         "/Generic_Windows8_StoreKey.pfx" + " " +
         packagePathStore + "/" + packagename + ".appx")


def buildPackages(root, prefix):
    buildAppxPackage(root, prefix, "ARM")
    signAppxPackage(root)
    
    buildAppxPackage(root, prefix, "x64")
    signAppxPackage(root)
    
    buildAppxPackage(root, prefix, "x86")
    signAppxPackage(root) 


def saveUpload():
    try:
        click("Gem")
    except:
        """
        SETAUTO(False)
        typeNum(34, TAB)
        press(ENTER)
        SETAUTO(True)
        """
        savebtn = Image("save_packages_icon.PNG")
        if savebtn.exists():
            click(savebtn)


def Menu_Pakker(root):
    try:
        wait_until(Text("Pakker").exists, timeout_secs=10)
    except:
        sleep(10)
        
    if Text("Pakker").exists() == False:
        SETAUTO(False)
        typeNum(19, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Pakker")

    log.printMessage("Going to Menu: Pakker")
    sleep(10)

    try:
        click("søg efter filer")
    except:
        SETAUTO(False)
        typeNum(26, TAB)
        press(ENTER)
        SETAUTO(True)

    openDialog()

    packagePath = root+metadataPath+"/AppxPackagesStore"
    
    appxfilexARM = getAPPXFile(packagePath, "ARM")
    appxfilex86 = getAPPXFile(packagePath, "x86")
    appxfilex64 = getAPPXFile(packagePath, "x64")

    if appxfilexARM == None:
        raise Exception("ERROR: No APPX - ARM file found...")

    if appxfilex86 == None:
        raise Exception("ERROR: No APPX - x86 file found...")

    if appxfilex64 == None:
        raise Exception("ERROR: No APPX - x64 file found...")

    
    try:
        write(packagePath)
        press(ENTER)
        sleep(5)
        press_and_hold(CTRL)
        click(appxfilexARM)
        click(appxfilex86)
        click(appxfilex64)
        release(CTRL)
        #click(appxfilexARM)
        #press(CTRL+'a')
        #press(CTRL+'a')
        log.printMessage("Adding packages...")
        press(ENTER)
    except:
        raise Exception("ERROR: Adding packages failed...")

    try:
        os.chdir(scripts_image_path)
        sleep(sleepTime)
        wait_until(Image("package_upload_done.PNG").exists, timeout_secs=120)
    except:
        pass
           
    press(PGDN)
    saveUpload()
    while(Image("bekraeft_navigation_packages.PNG").exists()):
        press(TAB)
        press(ENTER)
        sleep(10)
        saveUpload()

    log.printMessage("Successfully added APPX files") 
    sleep(10)

"**********END OF USE CASE 6: Pakker**********" 


"**********USE CASE 7: Beskrivelse**********" 
def getDescOfAppFromGoogleDocs():
    switchTo(googledocs)
    log.printMessage("Getting App Description from Google Docs...")
    SETAUTO(False)
    typeNum(8, TAB)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a',CTRL+'c')
    switchTo(menulistWindow)
    press(CTRL+'v')
    

def addImageToList(path):
    log.printMessage("Screenshot path " + path + " does exists!")
    os.chdir(path)
    imgList = []
    for image in os.listdir(os.getcwd()):
        if image.endswith(".png") and len(imgList) < maxImageUpload:
            log.printMessage("Adding image "+ image +" to list")
            imgList.append(image)

    return imgList
    

def findImages():
    log.printMessage("Finding screenshotpath..")
    try:
        changePathToDesktop()
        global screenshotPath
        screenshotPath = os.getcwd() + screenshotFolder + "/" + globapp

        if os.path.isdir(screenshotPath) == False:
            log.printMessage(screenshotPath + " does not exists")
            return None
        else:
            listOfImages = addImageToList(screenshotPath)
            return listOfImages
    except:
        log.printMessage("Exception: " + screenshotPath + " does not exists!!")
        return None
 

def uploadImages():
    openDialog() 
    imageList = findImages()
    if imageList == None:
        noImageMsg = "OBS OBS: No images found for " + appName + " !!!! Cannot submit app. Skipping to next!"
        log.printMessage(noImageMsg)
        raise Exception(noImageMsg)
    else:
        write(screenshotPath)
        press(ENTER)
        i = 0
        os.chdir(scripts_image_path)
        for image in imageList:
            log.printMessage("Uploading image..." + str(i+1))
            click(image)
            press(ENTER)
            try:
                wait_until(Image("skjermbilde_textfield.PNG").exists, timeout_secs=15)
            except:
                while(Image("skjermbilde_textfield.PNG").exists() == False):
                    press(ENTER)
                    openDialog()
                    click(image)
                    press(ENTER)
                    sleep(15)
            #sleep(10)
            press(TAB)
            write(imageCaption)
            i+=1
            if i == len(imageList):
                break
            else:
                typeNum(3, TAB)
                press(ENTER)
                openDialog()

        # Hvis der er ikke er screenshot af article view, så er der kun 4 screenshots. Ellers 5.
        if len(imageList) == 4:
            dynamic_tab = tab_nr_to_keywords+1
        else:
            dynamic_tab = tab_nr_to_keywords
            
        typeNum(dynamic_tab, TAB)


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
    
    log.printMessage("Seperating keywords...")

    """
    userhome = os.path.expanduser('~')
    desktop = userhome + '/Desktop/'
    os.chdir(desktop)
    """
    """
    OBS, hvis keywords delen køres seperat, erstat 'open' med 'file'
    og sæt 'tempfileName' som argument.
    """
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
    switchTo(menulistWindow)


def insertKeywords():
    log.printMessage("Inserting keywords...")
    i = 0
    while(i < len(keywords)):
        write(keywords[i])
        press(ESC, ESC) # Prevent website redirect.
        if i == maxAmountKeywords-1:
            break
        press(TAB)
        press(ESC)
        i += 1

    if len(keywords) < maxAmountKeywords:
        while(i < maxAmountKeywords-1):
            press(TAB)
            press(ESC)
            i += 1
        

def addKeywords():
    getKeywordsForSearchFromGoogleDocs()
    seperateKeywords()
    insertKeywords()


def addCopyright():
    typeNum(2, TAB)
    log.printMessage("Getting copyright text from Google Docs")
    switchTo(googledocs)
    press(TAB)
    press(ENTER)
    press(CTRL+'a',CTRL+'c')
    switchTo(menulistWindow)
    press(CTRL+'v')


def addWebURL():
    typeNum(10, TAB)
    log.printMessage("Getting Web URL from Google Docs")
    switchTo(googledocs)
    sleep(3)
    SETAUTO(False)
    typeNum(20, TAB)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a',CTRL+'c')
    switchTo(menulistWindow)
    press(CTRL+'v')


def addEmailSupport():
    typeNum(2, TAB)
    log.printMessage("Getting Email support text from Google Docs")
    switchTo(googledocs)
    sleep(3)
    SETAUTO(False)
    shiftTab(15)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a', CTRL+'c')
    switchTo(menulistWindow)
    press(CTRL+'v')


def addPrivaryURL():
    typeNum(2, TAB)
    log.printMessage("Getting privacy URL from Google Docs")
    switchTo(googledocs)
    press(TAB)
    press(ENTER)
    press(CTRL+'a', CTRL+'c')
    switchTo(menulistWindow)
    press(CTRL+'v')


def addInAppDescription():
    press(TAB)
    log.printMessage("Adding in-app description for each product...")
    i = 0
    
    while(i < 31):
        SETAUTO(False)
        press(TAB)
        write(InAppAlias.decode('utf-8'))
        SETAUTO(True)
        i+=1
        
    press(TAB)
    press(ENTER)
    
    
       
def Menu_Beskriv():
    try:
        wait_until(Text("Beskrivelse").exists, timeout_secs=15)
    except:
        sleep(15)
        
    if Text("Beskrivelse").exists() == False:
        SETAUTO(False)
        typeNum(20, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click("Beskrivelse")

    log.printMessage("Going to Menu: Beskrivelse")
    sleep(10)

    SETAUTO(False)
    typeNum(29, TAB)
    press(ENTER)
    sleep(5)

    typeNum(2, TAB)
    SETAUTO(True)

    getDescOfAppFromGoogleDocs()

    typeNum(9, TAB)
    sleep(3)
    shiftTab(3)
    press(ENTER)

    uploadImages()
    addKeywords()
    addCopyright()
    addWebURL()
    addEmailSupport()
    addPrivaryURL()
    addInAppDescription()

    sleep(10)

"**********END OF USE CASE 7: Beskrivelse**********"

"**********USE CASE 8: Bemaerkelse til testere**********" 
def Menu_BemTest():
    try:
        menuNoticeToTester = "Bemærkninger til testere".decode('utf-8')
        wait_until(Text(menuNoticeToTester).exists, timeout_secs=10)
    except:
        sleep(10)
        
    if Text(menuNoticeToTester).exists() == False:
        SETAUTO(False)
        typeNum(20, TAB)
        press(ENTER)
        SETAUTO(True)
    else:
        click(menuNoticeToTester)

    log.printMessage("Going to Menu: Bemaerkninger til testere")
    sleep(10)

    log.printMessage("Writing notice to testere: " + messageToTestere)
    write(messageToTestere)

    press(TAB)
    press(ENTER)

    sleep(10)

"**********END OF USE CASE 8: Bemaerkelse til testere**********" 


"**********SUBMIT AND START OVER**********" 
def submitApp():
    log.printMessage("Submitting app...")
    press(PGDN)
    try:
        click("Send til certificering")
    except:
        SETAUTO(False)
        typeNum(52, TAB)
        press(ENTER)
        SETAUTO(True)

    sleep(10)
    press(PGUP)


def resetGoogleDocs():
    log.printMessage("Resetting google docs...")
    switchTo(googledocs)
    SETAUTO(False)
    shiftTab(docsresetnum)
    press(ENTER)
    SETAUTO(True)
    press(CTRL+'a')
    write("SUBMITTED")
    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)

def restartApplications():
    SETAUTO(False)
    try:
        while Window("Google Chrome").exists():
            kill("Google Chrome")
    except:
        pass
    SETAUTO(True)

    start("Google Chrome", '-new-window', "https://appdev.microsoft.com/StorePortals/da-DK/Home/Index")
    sleep(15)
    
    start("Google Chrome", '-new-window', urldocs)
    sleep(15)

def restart():
    log.printMessage("RESTARTING......")
    os.chdir(rootdir)
    execfile('screenshot_log.py')
    
    switchTo(googledocs)
    sleep(5)

    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)

    searchInDocsForPrefix(visiolink_project_name)
    sleep(5)
    press(TAB)
    press(ENTER)
    press(CTRL+'a')
    write("FAILED")

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

    if Window("Lenovo Settings Power").exists():
        switch_to("Lenovo Settings Power")
        press(ALT+F4)

    """
    restartApplications()
    """


def printResults():
    log.printMessage("DONE - Successfully submittet app and in-app product for: " + appName)


def close():
    log.printMessage("Closing applications...")
    if Window(tempfileName).exists():
        switch_to(tempfileName)
        press(ALT+F4)

    if Window(tempfileName2).exists():
        switch_to(tempfileName2)
        press(ALT+F4)

    if Window("Microsoft Visual Studio").exists():
        kill("Microsoft Visual Studio")
        sleep(3)

    if Window("Lenovo Settings Power").exists():
        switch_to("Lenovo Settings Power")
        press(ALT+F4)

    try:
        switchTo(googledocs)
        os.chdir(scripts_image_path)
        wait_until(Image("changes_saved.PNG").exists, timeout_secs=10)
    except:
        pass

    SETAUTO(False)
    try:
        while Window("Google Chrome").exists():
            kill("Google Chrome")
    except:
        pass
    SETAUTO(True)


"**********END OF SUBMIT AND START OVER**********" 

 


"**********APP SUBMIT FLOW**********" 
def appmainflow(root,name, prefix):
    try:
        log.printMessage("*****************APP UPLOAD NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appName
        appName = getMetadataFromGoogleDocs(4, "App name")

        ClearPackageMap(root)
        buildPackages(root, prefix)

        if startUpFlag == True:
            switchTo(googledocs)
            start("Google Chrome", '-new-window', "https://appdev.microsoft.com/StorePortals/da-DK/Home/Index")
            try:
                wait_until(Window(dashboardWindow).exists, timeout_secs=120)
            except:
                pass
        
        goToDashboard()
        goToSendEnApp()
        goToMenu_Appnavn()
        Menu_Appnavn()
        Menu_SalgsDetal()
        Menu_Tjenester()
        Menu_AlderOgKlass()
        Menu_Krypto()

        Menu_Pakker(root)
        """
        global menulistWindow
        menulistWindow = appName.decode('iso-8859-1').encode('utf-8')
        switchTo(menulistWindow)
        """
        Menu_Beskriv()
        Menu_BemTest()

        #submitApp()

        printResults()
        resetGoogleDocs()
        log.printMessage("*****************END OF APP UPLOAD NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global startUpFlag
        startUpFlag = False
        global debugCounter
        debugCounter += 1
        global appCounter
        appCounter += 1
        sleep(5)
        # OBS OBS
        #debugWaitForNextSubmit()
        
    except:
        log.printMessage("Unexpected error: " + str(sys.exc_info()[1]))
        log.printMessage("ERROR - Failed submitting for " + name)
        restart()

"**********END OF APP SUBMIT FLOW**********" 


"**********MAIN**********" 
def foreach(prefix):
    for root, dirs, files in os.walk(rootdir+"/"+prefix):
        for name in files:
            if name.endswith(".sln"):
                global globapp
                globapp = prefix
                appmainflow(root,name,prefix)
                return 0


def submit(prefix):
    for dir in os.listdir(rootdir):
        prefix = prefix.replace("_W8", "")
        if prefix == dir:
            global visiolink_project_name
            visiolink_project_name = prefix + "_W8"
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

def appIsSubmitted():
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

    try:
        if(appIsSubmitted() == True):
            runList(subList)
            global startUpFlag
            startUpFlag = False

        resetMarkerInGoogleDocs()
    except:
        log.printMessage("Error: " + str(sys.exc_info()[1]))


global appCounter
appCounter = 1

global startUpFlag
startUpFlag = True

#Tjekker hele listen 3 gange for at tjekke at alt er gået godt.
def run():
    checkList()
    
    if appCounter == xappssubmits:
        log.printMessage("Maximum specified app submit reached..." + str(xappssubmits))
    log.printMessage("All apps are ready to be submitted or already submitted...")

run()
close()

"**********END OF MAIN**********" 

