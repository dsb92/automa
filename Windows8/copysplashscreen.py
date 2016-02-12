#!/usr/bin/python
# -*- coding: utf-8 -*-
import os, shutil, sys
import searchgraphics as graphics
from random import randint

def renamePng(fromname, toname, appname):
    ran_num = randint(123456, 7891011)
    os.rename(fromname, str(ran_num)+"_"+toname)


rootdir = os.path.dirname(os.path.realpath(__file__))

splashScreen = "b_620x300.scale-100.png"
os.chdir(rootdir)

appnames = file("appnames.txt")

for appname in appnames:
    graphicPath = graphics.getGraphicPath(rootdir,appname)
    
    if graphicPath == None:
        raise Exception("ERROR: Graphicpath not found.")

    os.chdir(graphicPath)
    source = os.listdir(graphicPath)
    for files in source:
        if files.endswith(".png") and "620x300" in files:
            destination = r"C:\Users\dab_000.LENOVO-PC.000\Desktop\splashscreenW8"
            shutil.copy(files,destination)
            os.chdir(destination)
            renamePng(graphics.splashScreen, splashScreen, appname)


"""
os.chdir(rootdir)
splashScreen = "b_620x300.scale-100.png"
appnames = file("appnames.txt")


for appname in appnames:
    graphicPath = graphics.getGraphicPath(rootdir,appname)
    
    if graphicPath == None:
        raise Exception("ERROR: Graphicpath not found.")

    os.chdir(graphicPath)
    source = os.listdir(graphicPath)
    for files in source:
        if splashScreen == files:
            os.remove(files)
    
"""
