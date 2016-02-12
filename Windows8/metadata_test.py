#!/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys, shutil
import xmlparser
from parameters import *
from subprocess import call
import log as log
import searchgraphics as graphics

os.chdir(r"C:\Users\dab_000.LENOVO-PC.000\Desktop")

w8appnames = file("w8appnames.txt")
testprefix = "haugesundsavis"
rootdir = os.path.dirname(os.path.realpath(__file__))
slnroot = r"C:\Users\David\Projekter\Device\Windows\Windows8"+"/"+testprefix+"\GenericReaderTemplate"
publisherid = "CN=5860D604-48AB-492F-9714-6351AE57306E"
prefix = testprefix

trackingaccount = "UA-50065389-21"
privacypolicy = r"https://sites.google.com/a/amedia.no/kons-si-app_privacy/"
aboutlocation = r"https://sites.google.com/a/amedia.no/kons-si-app_infohardangerfolkeblad/"
externalweb = "www.hardanger-folkeblad.no"
comscoreCustom = "10597850"
comscoreDomain = "visiolink.hardanger.eavis"
comscorePubSecret = "fc095c7b0f41b9a16176d7ca6841d974"
comscoreVirtual = "hardanger-folkeblad"
helppage = "www.hardanger-folkeblad.no"
colorcodetheme = "#0060A9"

def renamePng(fromname, toname):
    if os.path.exists(toname):
        os.remove(toname)
        os.rename(fromname, toname)
    else:
        os.rename(fromname, toname)

def copyGraphicsToProject(root):
    log.printMessage("Copying graphics to project...")
    storeLogo = "StoreLogo.scale-100.png"
    logo = "Logo.scale-100.png"
    smallLogo = "b_30x30.scale-100.png"
    wideLogo = "b_310x150.scale-100.png"
    splashScreen = "b_620x300.scale-100.png"

    graphicPath = graphics.getGraphicPath(rootdir,appName)
    
    if graphicPath == None:
        raise Exception("ERROR: Graphicpath not found.")

    os.chdir(graphicPath)
    source = os.listdir(graphicPath)
    for files in source:
        if files.endswith(".png"):
            destination = root+metadataPath
            shutil.copy(files,destination)

    os.chdir(root+metadataPath)

    renamePng(graphics.wideLogo, wideLogo)
    renamePng(graphics.storeLogo, storeLogo)
    renamePng(graphics.smallLogo, smallLogo)
    renamePng(graphics.logo, logo)
    renamePng(graphics.splashScreen, splashScreen)

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


x=1
for appname in w8appnames:
    global appName
    appName = appname
    xmlparser.xmlParse_SettingsXML(slnroot,
                         trackingaccount,
                         privacypolicy,
                         aboutlocation,
                         externalweb,
                         appname,
                         prefix,
                         comscoreCustom,
                         comscoreDomain,
                         comscorePubSecret,
                         comscoreVirtual,
                         helppage)
    xmlparser.xmlParse_SettingsXAML(slnroot, colorcodetheme, appname)
    xmlparser.xmlParse_PackageStoreAssoc(slnroot, publisherid, appname, prefix)
    xmlparser.xmlParse_PackageAppxManifest(slnroot, publisherid, appname, prefix)
    copyGraphicsToProject(slnroot)
    xml = xmlparser.xmlParse_server_xml_todayspaper(prefix)
    xmlparser.xmlParse_server_has_sections(xml)
    xmlparser.xmlParse_server_has_articles(xml)

    buildPackages(slnroot, prefix)
    
    print str(x)
    x+=1
