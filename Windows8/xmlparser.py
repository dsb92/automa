#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import urllib
import xml.etree.ElementTree as ET
from datetime import datetime
current_date = datetime.now().date().isoformat()

path = "/WindowsGenericReader/WindowsGenericReader"
settingsXAML = "Settings.xaml"
settingsXML = "Settings.xml"
packageAppxManifest = "Package.appxmanifest"
packageStoreAss = "Package.StoreAssociation.xml"
culturecode = "nb-NO"

"""#UNIT TEST AF SETTINGS.XML
slnroot = r"C:\Users\David\Projekter\aMediaTest\Device\NytSetup\WindowsPhone\drammenstidende"
trackingaccount = "UA-50065389-5"
privacypolicy = "https://sites.google.com/a/amedia.no/kons-si-app_privacy/"
aboutlocation = "https://sites.google.com/a/amedia.no/kons-si-app_infodrammenstidende/"
externalweb = "www.dt.no"
appname = "Drammens Tidende"
prefix = "drammenstidende"
comscoreCustom = "10597850"
comscoreDomain = "visiolink.drammenstidende.eavis"
comscorePubSecret = "fc095c7b0f41b9a16176d7ca6841d974"
comscoreVirtual = "dt"
helppage = "http://www.visiolink.com"

xmlParse_SettingsXML(slnroot,
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

"""     

def xmlParse_SettingsXML(slnroot,
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
                         helppage):
    os.chdir(slnroot+path)
    tree = ET.parse(settingsXML)

    root = tree.getroot()

    for child in root:
        if child.tag == "appname":
            child.text = appname.decode('iso-8859-1')
        
        if child.tag == "customer":
            child.text = prefix

        if child.tag == "privacyurl":
            child.text = privacypolicy

        if child.tag == "abouturl":
            child.text = aboutlocation

        if child.tag == "helpurl":
            if "http" in helppage:
                child.text = helppage
            else:
                child.text = "http://" + helppage

        if child.tag == "trackingaccount":
            child.text = trackingaccount

        if child.tag == "externalwebview":
            if "http" in externalweb:
                child.text = externalweb
            else:
                child.text = "http://" + externalweb

        if child.tag == "comscorec2":
            child.text = comscoreCustom

        if child.tag == "comscorepublishersecret":
            child.text = comscorePubSecret

        if child.tag == "comscoreappname":
            child.text = comscoreDomain

        if child.tag == "comscoresitenamecode":
            child.text = comscoreVirtual

        tree.write(settingsXML, "utf-8", True)
    

"""
# TIL 'UNIT' test af SETTINGS.XAML.
os.chdir(r"C:\Users\David\Projekter\aMediaTest\Device\NytSetup\WindowsPhone\ringsakerblad")
ET.register_namespace('',"http://schemas.microsoft.com/winfx/2006/xaml/presentation")
ET.register_namespace('x',"http://schemas.microsoft.com/winfx/2006/xaml")
ET.register_namespace('system', "clr-namespace:System;assembly=mscorlib")
ET.register_namespace('toolkit', "clr-namespace:Microsoft.Phone.Controls;assembly=Microsoft.Phone.Controls.Toolkit")
tree = ET.parse("Settings.xaml")

root = tree.getroot()
xmlns="{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"
x="{http://schemas.microsoft.com/winfx/2006/xaml}"
attcolor = "CurrentColor"
for color in root.iter(xmlns+"Color"):
        if attcolor in color.attrib.get(x+"Key"):
            color.text = "#F0F0F0F0"
            tree.write("Settings.xaml", "utf-8", True)
            break

with open(settingsXAML, 'r') as file:
        data = file.readlines()

os.chdir(r"C:\Users\dab_000.LENOVO-PC.000\Desktop")
datafile = open("test.txt", 'r')
for line in datafile:
    appname = line
appname = str(appname.decode('iso-8859-1').encode('utf-8'))
i=0
x = '<system:String x:Key="AppName">'+appname+'</system:String>'
z = 'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
y = 'xmlns:toolkit="clr-namespace:Microsoft.Phone.Controls;assembly=Microsoft.Phone.Controls.Toolkit"'
for d in data:
    if z in d:
        data[i] = data[i].replace(z, z+" "+y)
    if "AppName" in d:
        data[i] = "\t\t\t"+x+"\n"
        break
    i+=1


with open(settingsXAML, 'w') as file:
    file.writelines(data)
"""

def xmlParse_SettingsXAML(slnroot, colorcodetheme, appname):
    os.chdir(slnroot+path)
    ET.register_namespace('',"http://schemas.microsoft.com/winfx/2006/xaml/presentation")
    ET.register_namespace('x',"http://schemas.microsoft.com/winfx/2006/xaml")
    #ET.register_namespace('system', "clr-namespace:System;assembly=mscorlib")
    #ET.register_namespace('toolkit', "clr-namespace:Microsoft.Phone.Controls;assembly=Microsoft.Phone.Controls.Toolkit")
    
    tree = ET.parse(settingsXAML)

    root = tree.getroot()
    xmlns="{http://schemas.microsoft.com/winfx/2006/xaml/presentation}"
    x="{http://schemas.microsoft.com/winfx/2006/xaml}"
    attcolor = "CurrentColor"
    for color in root.iter(xmlns+"Color"):
        if attcolor in color.attrib.get(x+"Key"):
            color.text = colorcodetheme
            tree.write(settingsXAML, "utf-8", True)
            break

    with open(settingsXAML, 'r') as file:
        data = file.readlines()

    appname = str(appname.decode('iso-8859-1').encode('utf-8'))
    appname = appname.rstrip()
    i=0
    x = '<x:String x:Key="AppName">'+appname+'</x:String>'
    #z = 'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'
    #y = 'xmlns:toolkit="clr-namespace:Microsoft.Phone.Controls;assembly=Microsoft.Phone.Controls.Toolkit"'
    for d in data:
        #if z in d:
        #    data[i] = data[i].replace(z, z+" "+y)
        if "AppName" in d:
            data[i] = "\t\t\t"+x+"\n"
            break
        i+=1

    with open(settingsXAML, 'w') as file:
        file.writelines(data)


def xmlParse_PackageAppxManifest(slnroot, publisherid, appname, prefix, colorcodesplash, colorcodehomeicon):
    os.chdir(slnroot+path)
    ET.register_namespace('',"http://schemas.microsoft.com/appx/2010/manifest")
    tree = ET.parse(packageAppxManifest)

    root = tree.getroot()
    xmlns="{http://schemas.microsoft.com/appx/2010/manifest}"
    
    for child in root:
        if child.tag == xmlns+"Identity":
            child.set('Name', packageId)
            child.set('Publisher', publisherid)
            child.set('Version', "1.0.0.0")
            break
        
    for properties in root.iter(xmlns+'DisplayName'):
        properties.text = appname.decode('iso-8859-1')
        break

    for viselement in root.iter(xmlns+'VisualElements'):
        viselement.set('BackgroundColor', colorcodehomeicon)
        viselement.set('DisplayName', appname.decode('iso-8859-1'))
        viselement.set('Description', appname.decode('iso-8859-1'))
        break

    for splashscreen in root.iter(xmlns+'SplashScreen'):
        splashscreen.set('BackgroundColor', colorcodesplash)
        break

    tree.write(packageAppxManifest, "utf-8", True)

"""UNIT TEST AF PackageAppxManifest

root = r"C:\Users\David\Projekter\Device\Windows\Windows8\haugesundsavis\GenericReaderTemplate"
publisherid = "CN=5860D604-48AB-492F-9714-6351AE57306E"
appname = "Øyene Digital Utgave"
prefix = "oyene"

xmlParse_PackageAppxManifest(root, publisherid, appname, prefix)
"""

def xmlParse_PackageStoreAssoc(slnroot, publisherid, appname, prefix):
    os.chdir(slnroot+path)
    ET.register_namespace('',"http://schemas.microsoft.com/appx/2010/storeassociation")
    tree = ET.parse(packageStoreAss)

    root = tree.getroot()
    xmlns="{http://schemas.microsoft.com/appx/2010/storeassociation}"

    for child in root:
        if child.tag == xmlns+"Publisher":
            child.text = publisherid

        if child.tag == xmlns+"PublisherDisplayName":
            child.text = "Amedia AS"
            break

    shortappname = appname.decode('iso-8859-1')
    shortappname = shortappname.replace(" ", "")
    shortappname = shortappname.replace("/", "")
    shortappname = shortappname.rstrip()
    #print shortappname
    if shortappname == "HamarDagblad":
        shortappname = "AmediaAS." + "61454361C04EE"
    else:
        shortappname = "AmediaAS." + shortappname
    
    i=0
    while(i<len(shortappname)):
        if shortappname[i].encode('utf-8') == 'Ø':
            shortappname = shortappname.encode('utf-8').replace("Ø", "")
        elif shortappname[i].encode('utf-8') == 'ø':
            shortappname = shortappname.encode('utf-8').replace("ø", "")
            
        elif shortappname[i].encode('utf-8') == 'Å':
            shortappname = shortappname.encode('utf-8').replace("Å", "")

        elif shortappname[i].encode('utf-8') == 'å':
            shortappname = shortappname.encode('utf-8').replace("å", "")
        i+=1

    shortappname = shortappname.rstrip()
    #print shortappname
    
    for pkgId in root.iter(xmlns+'MainPackageIdentityName'):
        if shortappname == pkgId.text:
            global packageId
            packageId = pkgId.text
            #print "EQUAL " + packageId
            break

    for pkgId in root.iter(xmlns+'MainPackageIdentityName'):
        pkgId.text = packageId
        break

    for reservedName in root.iter(xmlns+'ReservedName'):
        reservedName.text = appname.decode('iso-8859-1')
        break

    tree.write(packageStoreAss, "utf-8", True)


"""#UNIT TEST AF PackageStore Ass :) 

slnroot = r"C:\Users\David\Projekter\Device\Windows\Windows8\haugesundsavis\GenericReaderTemplate"
publisherid = "CN=5860D604-48AB-492F-9714-6351AE57306E"
appname = "Øyene Digital Utgave"
prefix = "oyene"


xmlParse_PackageStoreAssoc(slnroot, publisherid, appname, prefix)
"""
            
            
def getVersionAppx(slnroot):
    os.chdir(slnroot+"/WindowsGenericReader/WindowsGenericReader")
    ET.register_namespace('',"http://schemas.microsoft.com/appx/2010/manifest")
    xmlns="{http://schemas.microsoft.com/appx/2010/manifest}"

    tree = ET.parse(packageAppxManifest)
    root = tree.getroot()

    for child in root:
        if child.tag == xmlns+"Identity":
            return child.get('Version')


def getTodaysCatalog(f):
    tree = ET.parse(f)
    root = tree.getroot()

    for child in root:
        for c in child:
            if c.tag == "catalog":
                return c.text
    

def xmlParse_server_xml_todayspaper(prefix):
    paper = "http://device.e-pages.dk/content/available4.php?customer="+prefix
    f = urllib.urlopen(paper)
    catalog = str(getTodaysCatalog(f))
    
    todayspaper = "http://device.e-pages.dk/content/default4.php?customer="+prefix+"&catalog="+catalog
    h = urllib.urlopen(todayspaper)

    tree = ET.parse(h)
    xml = tree.getroot()

    return xml
    

def xmlParse_server_has_sections(xml):
    for child in xml.iter("sections"):
        if len(child) > 1:
            return True
        else:
            return False


def xmlParse_server_has_articles(xml):
    if any("articles" in child.tag for child in xml):
        return True
    else:
        return False


""" UNIT TEST AF SERVER SIDE SCRIPTING
prefix = "oyene"
xml = xmlParse_server_xml_todayspaper(prefix)
xmlParse_server_has_sections(xml)
xmlParse_server_has_articles(xml)
"""

    
    
    
    
    
    
