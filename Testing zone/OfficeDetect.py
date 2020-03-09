#Office detection

import os 

'''
Description about this file:
==============================
This script cheking MS Office version on your PC, if you dont have MS Office on you PC, script returning "You dont have Office on your PC :c"
==============================

'''

#Getting path to ProgramFiles folder
progFilesFolder = os.environ["PROGRAMFILES"]

#Office versions variables
office10 = progFilesFolder+r"\Microsoft Office\Office14\ospp.vbs"
office13 = progFilesFolder+r"\Microsoft Office\Office15\ospp.vbs"
office = progFilesFolder+r"\Microsoft Office\Office16\ospp.vbs"

#Office version detection function
def checkOfficeVer():
    if os.path.exists(office10):
        print("Your Office version: Office 2010!")
    elif os.path.exists(office13):
        print("Your Office version: Office 2013!")
    elif os.path.exists(office):
        print("Your Office version: Office 2016, 2019 or 365!")
    else:
        print("You dont have Office on your PC :c")


'''
Why office version 2016, 2019 or 365? :
=================
Because Office 2016-365 they have no differences in paths and are difficult to distinguish.
=================
'''

checkOfficeVer()

input()




