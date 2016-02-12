from subprocess import call

signExe = r"C:\Program Files (x86)\Windows Kits\8.1\bin\x64\signtool.exe"

rootdir = r"C:\Users\David\Projekter\Device\Windows\Windows8\oyene\GenericReaderTemplate"

appx = r"C:\Users\David\Projekter\Device\Windows\Windows8\oyene\GenericReaderTemplate\WindowsGenericReader\WindowsGenericReader\Test_1.0.0._AnyCPU.appx"


call(signExe + " sign /fd SHA256 /a /f "+rootdir+"/WindowsGenericReader/WindowsGenericReader/Generic_Windows8_StoreKey.pfx" + " " + appx)
