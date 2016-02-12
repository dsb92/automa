#!/usr/bin/python
# -*- coding: utf-8 -*-
import os

rootdir = os.path.dirname(os.path.realpath(__file__))

os.chdir(rootdir)
datafile = file('prefixes.txt')

os.chdir(r"D:\amedia\Windows8")
for line in datafile:
    line = line.rstrip()
    os.mkdir(line)
    
                
