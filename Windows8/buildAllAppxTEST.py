import os, shutil
from subprocess import call
import xmlparser
from parameters import *
from os.path import expanduser

rootdir = os.path.dirname(os.path.realpath(__file__))
mypath = r"C:\Users\dab_000.LENOVO-PC.001\Desktop\packages"
appSubmits = 67

def removeIfAPPXExists(root, buildconfig):
    os.chdir(packagePathStore)
    source = os.listdir(os.getcwd())
    
    for files in source:
        if files.endswith(".appx"):
            version_current_date = buildconfig+"_"+str(appxversion)+"_"+str(xmlparser.current_date)
            if version_current_date in files:
                os.remove(files)
                break

def getPackageName(prefix, buildconfig, version):
    return prefix+"_"+buildconfig+"_"+version+"_"+str(xmlparser.current_date)

def buildAppxPackage(root, prefix, buildconfig):
    global appxversion
    appxversion = xmlparser.getVersionAppx(root)

    global packagename
    packagename = getPackageName(prefix, buildconfig, appxversion)

    global packagePathStore
    packagePathStore = mypath + "/" + prefix

    if os.path.exists(packagePathStore) == False:
        os.mkdir(packagePathStore)
    else:
        removeIfAPPXExists(root, buildconfig)

    package_map_path_app = mypath + "/" + prefix+"/package_maps"
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


def foreach(prefix):
    for root, dirs, files in os.walk(rootdir+"/"+prefix):
        for name in files:
            if name.endswith(".sln"):
                buildPackages(root, prefix)
                return 0

appxCounter = 0
buildErrorList = []
buildList = []
os.chdir(rootdir)
datafile = file("prefixes.txt")

def getAppxCount(prefix):
    os.chdir(mypath + "/" + prefix)
    source = os.listdir(os.getcwd())

    for appx in source:
        if appx.endswith(".appx"):
            global appxCounter
            appxCounter += 1

    return appxCounter

for line in datafile:
    prefix = line.rstrip()
    foreach(prefix)
    if getAppxCount(prefix) == 3:
        buildList.append(prefix)
        appxCounter = 0
    else:
        buildErrorList.append(prefix)
        appxCounter = 0
    os.chdir(rootdir)

print str(len(buildList)) + "/" + str(appSubmits) + " Build Succeeded!"

if len(buildList) != appSubmits:
    print "FAILED:"
    for errbuild in buildErrorList:
        print errbuild
    


