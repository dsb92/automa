import os
from parameters import *

rootdir = os.path.dirname(os.path.realpath(__file__))

i = 1

def foreach(d):
    for root,dirs, files in os.walk(rootdir+"/"+d):
        for name in files:
            if name.endswith(".sln"):
                path = root + metadataPath + "/AppxPackagesStore"
                
                if os.path.exists(path):
                    os.chdir(path)
                    source = os.listdir(os.getcwd())
                    for s in source:
                        if s.endswith(".appx"):
                            os.remove(s)
                            print s + " " + str(i)
                            print root
                            global i
                            i += 1
                return 0
                

def run():
    for dir in os.listdir(rootdir):
        foreach(dir)


run()


