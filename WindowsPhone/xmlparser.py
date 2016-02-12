#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import urllib
import xml.etree.ElementTree as ET
from datetime import datetime
current_date = datetime.now().date().isoformat()

settingsXAML = "Settings.xaml"
settingsXML = "Settings.xml"
wmappmanifestXML = "WMAppManifest.xml"
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
    os.chdir(slnroot)
    tree = ET.parse(settingsXML)

    root = tree.getroot()

    for child in root:
        if child.tag == "appname":
            child.text = appname.decode('iso-8859-1')
        
        if child.tag == "customer":
            child.text = prefix

        if child.tag == "privacypolicy":
            child.text = privacypolicy

        if child.tag == "aboutlocation":
            child.text = aboutlocation

        if child.tag == "helppage":
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
    os.chdir(slnroot)
    ET.register_namespace('',"http://schemas.microsoft.com/winfx/2006/xaml/presentation")
    ET.register_namespace('x',"http://schemas.microsoft.com/winfx/2006/xaml")
    ET.register_namespace('system', "clr-namespace:System;assembly=mscorlib")
    ET.register_namespace('toolkit', "clr-namespace:Microsoft.Phone.Controls;assembly=Microsoft.Phone.Controls.Toolkit")
    
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
# TIL 'UNIT' test af WMAppManifest.xml
os.chdir(r"C:\Users\David\Projekter\aMediaTest\Device\NytSetup\WindowsPhone\Base\Properties")
#ET.register_namespace('',"http://schemas.microsoft.com/windowsphone/2012/deployment")
tree = ET.parse("WMAppManifest.xml")
root = tree.getroot()

for child in root:
    if child.tag == "DefaultLanguage":
        child.set('code', 'nb-NO')
        break

for lang in root.iter('Language'):
    lang.set('code', 'nb-NO')
    break


for child in root:
    if child.tag == "App":
        child.set('ProductID', 'test')
        child.set('Publisher', 'test')
        child.set('PublisherID', 'test')
        child.set('Title', 'test')
        break

tree.write("WMAppManifest.xml", "utf-8", True)
"""

def xmlParse_WMAppManifest(slnroot, productid, publisher, appname, publisherid):
    os.chdir(slnroot+"/Properties")
    ET.register_namespace('x',"http://schemas.microsoft.com/windowsphone/2012/deployment")
    tree = ET.parse(wmappmanifestXML)
    root = tree.getroot()

    for child in root:
        if child.tag == "DefaultLanguage":
            child.set('code', culturecode)
            break

    for lang in root.iter('Language'):
        lang.set('code', culturecode)
        break

    for iconpath in root.iter('IconPath'):
        iconpath.text = "Winnie\..\Assets\ApplicationIcon.png"
        break

    for smalltile in root.iter('SmallImageURI'):
        smalltile.text = "Winnie\..\Assets\Tiles\FlipCycleTileSmall.png"
        break

    for medtile in root.iter('BackgroundImageURI'):
        medtile.text = "Winnie\..\Assets\Tiles\FlipCycleTileMedium.png"

    for title in root.iter('Title'):
        title.text = appname.decode('iso-8859-1')

    for child in root:
        if child.tag == "App":
            child.set('Author', publisher + " author")
            child.set('ProductID', "{"+productid+"}")
            #child.set('ProductID', "{415e3737-82d6-4759-9bde-f6b693a64b03}")
            #child.set('Publisher', publisher)
            child.set('PublisherID', "{"+publisherid+"}")
            child.set('Title', appname.decode('iso-8859-1'))
            child.set('Version', "1.0.0.3")
            break

    tree.write(wmappmanifestXML, "utf-8", True)
    

""" TIL 'UNIT' test af CSPROJ for SplashScreenImage.jpg    
os.chdir(r"C:\Users\David\Projekter\aMediaTest\Device\NytSetup\WindowsPhone\Base")
xmlns="{http://schemas.microsoft.com/developer/msbuild/2003}"
cproj = "VISIOLINK_PROJECT_NAME.csproj"
ET.register_namespace('',"http://schemas.microsoft.com/developer/msbuild/2003")
tree = ET.parse(cproj)
xaptag = "XapFilename"
root = tree.getroot()

for content in root.iter(xmlns+"PropertyGroup"):
    for c in content:
        if xaptag in c.tag:
            root, ext = os.path.splitext(c.text)
            root = root+"_1.0.02"+"_27102014"
            c.text = root+ext
            tree.write(cproj, "utf-8", True)
"""

def getVersionOfXAP(slnroot):
    os.chdir(slnroot+"/Properties")
    tree = ET.parse(wmappmanifestXML)
    root = tree.getroot()

    for child in root:
        if child.tag == "App":
            return child.get('Version')


def xmlParse_CSPROJ(slnroot, version):
    versionNum = version
    os.chdir(slnroot)
    xmlns="{http://schemas.microsoft.com/developer/msbuild/2003}"

    for filename in os.listdir(slnroot):
        if filename.endswith(".csproj"):
            csproj = filename

    if csproj == None:
        raise Exception("No project file '.csproj' found in root of solution project")

    ET.register_namespace('',"http://schemas.microsoft.com/developer/msbuild/2003")
    tree = ET.parse(csproj)
    xaptag = "XapFilename"
    root = tree.getroot()

    for content in root.iter(xmlns+"PropertyGroup"):
        for c in content:
            if xaptag in c.tag:
                root, ext = os.path.splitext(c.text)
                root = root+"_"+versionNum+"_"+str(current_date)
                c.text = root+ext
                tree.write(csproj, "utf-8", True)
                return c.text



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

    
    
    
    
    
    
