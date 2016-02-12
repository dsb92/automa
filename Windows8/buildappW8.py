#!/usr/bin/python
# -*- coding: utf-8 -*-

from automa.api import *
from win32api import GetSystemMetrics
from subprocess import call
from random import randint
from time import sleep
import os, sys, shutil
import log as log
import xmlparser
import searchgraphics as graphics
from parameters import *

urldocs = "https://docs.google.com/spreadsheets/d/1e0sswHJo91jPbLfAdzTrC2dfFDbpGP6m8f81aMtP_2A/edit#gid=614091273"

"Directory from where this script is located"
rootdir = os.path.dirname(os.path.realpath(__file__))
os.chdir(rootdir)

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

   
def resetToAppNameInGoogleDocs(num):
    switchTo(googledocs)
    SETAUTO(False)
    shiftTab(num)
    SETAUTO(True)


def editSettingsXML(root, prefix):
    log.printMessage("Editing Settings.xml file...")

    newTrackingAccount = getMetadataFromGoogleDocs(5, "Tracking account")
    newPrivacyPolicy = getMetadataFromGoogleDocs(11, "Privacy policy")
    #newAboutLocation = getMetadataFromGoogleDocs(9, "About location")
    newAboutLocation = "http://device.e-pages.dk/content/amedia/webview.php?view=info&customer="+prefix
    #newWebURL = getMetadataFromGoogleDocs(3, "External Web URL")
    newWebURL = "http://device.e-pages.dk/content/amedia/webview.php?view=home&customer="+prefix+"&platform=windows&device_type=tablet"
    newComscoreCustom = getMetadataFromGoogleDocs(19, "Comscore customer C2")
    newComscoreDomain = getMetadataFromGoogleDocs(1, "Comscore domain/app name")
    newComscorePubSecret = getMetadataFromGoogleDocs(1, "Comscore publisher secret")
    newComscoreVirtual = getMetadataFromGoogleDocs(1, "Comscore virtual")

    #newHelppage = getMetadataFromGoogleDocs(1, "Help page")
    newHelppage = "http://device.e-pages.dk/content/amedia/webview.php?view=help&customer="+prefix

    #Reset to 'About location' column
    resetToAppNameInGoogleDocs(13)

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


def editSettingsXAML(root, prefix):
    log.printMessage("Editing Settings.xaml file...")

    newColorCodeTheme = getMetadataFromGoogleDocs(1, "Color code theme")
    xmlparser.xmlParse_SettingsXAML(root, newColorCodeTheme, appName)


def editPackageAppxManifest(root, prefix):
    log.printMessage("Editing PackageAppxManifest XML file")

    newColorCodeSplashBg = getMetadataFromGoogleDocs(1, "Color code background splash")
    newColorCodeHomeIcon = getMetadataFromGoogleDocs(1, "Color code home icon")
    xmlparser.xmlParse_PackageAppxManifest(root, publisherid, appName, prefix, newColorCodeSplashBg, newColorCodeHomeIcon)
    resetToAppNameInGoogleDocs(28)


def editPackageStoreAssoc(root, prefix):
    log.printMessage("Editing Package.StoreAssociation XML file")

    xmlparser.xmlParse_PackageStoreAssoc(root, publisherid, appName, prefix)


def renamePng(fromname, toname):
    if os.path.exists(toname):
        os.remove(toname)
        os.rename(fromname, toname)
    else:
        os.rename(fromname, toname)


def graphicOperation(source, destination, operation):
    for files in source:
        if files.endswith(".png"):
            if operation == "remove":
                os.remove(files)
            elif operation == "copy":
                shutil.copy(files, destination)
            else:
                raise Exception("ERROR: Please specify graphic operation...")

def copyGraphicsToProject(root):
    log.printMessage("Copying graphics to project...")
    dummySmallLogo = "tile_30-30.png"
    dummyLogo = "tile_150-150.png"
    dummyWideLogo = "tile_310-150.png"
    
    storeLogo = "StoreLogo.scale-100.png"
    logo = "Logo.scale-100.png"
    smallLogo = "b_30x30.scale-100.png"
    wideLogo = "b_310x150.scale-100.png"
    splashScreen = "b_620x300.scale-100.png"

    graphicPath = graphics.getGraphicPath(rootdir,appName)
    
    if graphicPath == None:
        raise Exception("ERROR: Graphicpath not found.")

    destination = root+metadataPath+"/images"
    os.chdir(destination)
    source = os.listdir(destination)

    #Remove existing image in ../images folder
    graphicOperation(source, destination, "remove")

    os.chdir(graphicPath)
    source = os.listdir(graphicPath)

    #Copy graphics to ../images
    graphicOperation(source, destination, "copy")

    destination = root+metadataPath
    os.chdir(graphicPath)
    source = os.listdir(graphicPath)

    #Copy graphics to ../WindowsGenericReader/WindowsGenericReader
    graphicOperation(source, destination, "copy")

    os.chdir(root+metadataPath)

    #Rename all images to graphics added to project
    renamePng(graphics.wideLogo, wideLogo)
    renamePng(graphics.storeLogo, storeLogo)
    renamePng(graphics.smallLogo, smallLogo)
    renamePng(graphics.logo, logo)
    renamePng(graphics.splashScreen, splashScreen)

    os.chdir(graphicPath)
    #Copy graphics to ../WindowsGenericReader/WindowsGenericReader AGAIN
    for files in source:
        if files.endswith(".png") and ("image_30x30" in files or "image_150x150" in files or "image_310x150" in files):
            shutil.copy(files, destination)

    
    os.chdir(root+metadataPath)
    #Rename graphics to dummy tiles not used in project 
    renamePng(graphics.smallLogo, dummySmallLogo)
    renamePng(graphics.logo, dummyLogo)
    renamePng(graphics.wideLogo, dummyWideLogo)


def resetGoogleDocs():
    log.printMessage("Resetting google docs...")
    switchTo(googledocs)
    press(TAB)
    press(ENTER)
    press(CTRL+'a')
    write("BUILT")
    press(CTRL+'h')
    write(prefixtitelcolumn)
    press(ENTER)
    SETAUTO(False)
    typeNum(11, TAB)
    press(ENTER)
    SETAUTO(True)


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
    switch_to(tempfileName)
    press(CTRL+'a',DEL,CTRL+'s')
    global startUpFlag
    startUpFlag = False

    if Window("Gem som").exists():
        switch_to("Gem som")
        press(ALT+F4)

    SETAUTO(False)
    try:
        while Window("Google Chrome").exists():
            kill("Google Chrome")
    except:
        pass
    SETAUTO(True)

    start("Google Chrome", '-new-window', urldocs)
    sleep(15)
    
    
def printResults():
    log.printMessage("DONE - Successfully submittet app and in-app product for: " + appName)


def close():    
    if Window("Microsoft Visual Studio").exists():
        kill("Microsoft Visual Studio")
        sleep(3)


"**********END OF SUBMIT AND START OVER**********" 

def beginMetaDataExtraction(root, prefix):
    editSettingsXML(root, prefix)
    editSettingsXAML(root, prefix)
    editPackageStoreAssoc(root, prefix)
    editPackageAppxManifest(root, prefix)
    copyGraphicsToProject(root)


"""
def buildApp(root, prefix):
    buildmsg = "Building app for... " + prefix
    devenvPath = r"C:\Program Files (x86)\Common7\IDE\devenv.exe"

    log.printMessage(buildmsg)
    call("echo " + buildmsg, shell=True)

    call(devenvPath + " " + root + " " + "/Build Release")



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

    i=0
    for d in data:
        data[i] = data[i].replace("haugesundsavis", prefix)
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
"""  

"**********APP BUILD FLOW**********" 
def appmainflow(root,name, prefix):
    try:
        log.printMessage("*****************APP BUILD NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appName
        appName = getMetadataFromGoogleDocs(4, "App name")

        beginMetaDataExtraction(root, prefix)
        #buildApp(root, prefix)
        #buildPackages(root, prefix)

        global startUpFlag
        startUpFlag = False
        log.printMessage("*****************END OF APP BUILD NR: " + str(appCounter) + "/" + str(xappssubmits)+ " for " + name + "*****************")
        global appCounter
        appCounter += 1
        searchInDocsForPrefix(visiolink_project_name)
        sleep(5)
        resetGoogleDocs()
    except:
        log.printMessage("Unexpected error: " + str(sys.exc_info()[1]))
        log.printMessage("ERROR - Failed submitting for " + name)
        restart()
"**********END OF APP BUILD FLOW**********" 


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
        if "NO" in sub:
            prefix = sub.replace("_NO", "")
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
    if any("NO" in sub for sub in subList):
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
        
    if(appIsSubmitted() == True):
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
    checkList()
    
    if appCounter == xappssubmits:
        log.printMessage("Maximum specified app build reached..." + str(xappssubmits))
    log.printMessage("All apps are built...")

run()
close()

"**********END OF MAIN**********" 

