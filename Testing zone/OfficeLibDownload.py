import os
import wget

'''
About this file:
==============================
This script loads external libraries for Office 2010 and Office 2013. 
They are needed for activation through the KMS server, 
because Office 2010 and 2013 do not have these libraries by default.
==============================
'''

def download():
    print("Downloading Office libs...")
    lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2013.zip?raw=true"
    lib10 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2010.zip?raw=true"
    wget.download(lib13)
    wget.download(lib10)

download()

input()
