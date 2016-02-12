#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
from random import randint
import sys
import time
import Image
import ImageGrab
#INSTALL PIL/Image MODULE HERE: http://www.pythonware.com/products/pil/
#---------------------------------------------------------
#User Settings:
userhome = os.path.expanduser('~')
desktop = userhome + '/Desktop/'
os.chdir(desktop)
SaveDirectory=os.getcwd() + "/screenshot_log"
if os.path.exists(SaveDirectory) == False:
    os.mkdir(SaveDirectory)

ImageEditorPath=r'C:\WINDOWS\system32\mspaint.exe'
#Here is another example:
#ImageEditorPath=r'C:\Program Files\IrfanView\i_view32.exe'
#---------------------------------------------------------
rand_num = randint(123456, 789101)
img=ImageGrab.grab()
saveas=os.path.join(SaveDirectory,'error_'+str(rand_num)+'.png')
img.save(saveas)
#editorstring='""%s" "%s"'% (ImageEditorPath,saveas) #Just for Windows right now?
#Notice the first leading " above? This is the bug in python that no one will admit...
#os.system(editorstring)
