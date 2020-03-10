import os
import wget
import time
import zipfile

'''
About this file:
==============================
This script loads external libraries for Office 2010 and Office 2013. 
They are needed for activation through the GMS server, 
because Office 2010 and 2013 do not have these libraries by default.
==============================
'''

def download():
    try:
        os.mkdir("C:/libraries")
        os.mkdir("C:/libraries/office13")
        input("Press any key to start...")
        print("\nЗагрузка внешней библеотеки для активаци...")
        lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2013.zip?raw=true"
        lib10 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2010.zip?raw=true"
        out = "C:/libraries/"
        wget.download(lib13, out=out)
        wget.download(lib10, out=out)
        libs = zipfile.ZipFile('C:/libraries/office2013.zip')
        libs.extractall('C:/libraries/office13/')
        time.sleep(3)
        print("\nЗавершено!")
    except OSError:
        input("Press any key to start...")
        print("\nЗагрузка внешней библеотеки для активаци...")
        lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2013.zip?raw=true"
        lib10 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2010.zip?raw=true"
        out = "C:/libraries/"
        wget.download(lib13, out=out)
        wget.download(lib10, out=out)
        libs = zipfile.ZipFile('C:/libraries/office2013.zip')
        libs.extractall('C:/libraries/office13/')
        time.sleep(3)
        print("\nЗавершено!")


download()

input()
