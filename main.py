import platform
import os
import colorama
from colorama import Back
import activ as act
from strings import Windows10, Windows8, Windows81, KMSdata
import locale
colorama.init()

#Auto translate
loc = locale.getdefaultlocale()[0]
getlang = loc[0:2]
if getlang == "ru":
    from strings import langRU as lang
else:
    from strings import langEN as lang


#Welcome message
print(":===================FixAct 1.0===================:")
print(lang.welcome)
print(lang.mainSelect)
select = input(lang.InputSelect)

#Variables
sysVer = platform.version()
sysNum = platform.win32_ver()[0]
sysEdition = platform.win32_edition()
sysName = sysNum + sysEdition
#===============================

#Activated Product Selection
def softSelect():
    os.system("cls")
    print(lang.softChoose)
    print("1) Windows\n2) Office\n ")
    softSelected = input(lang.softSelect)
    if softSelected == "1":
        WinActConfirm()
    if softSelected == "2":
        print("\nOffice activation do not work now :c")
    else:
        print(lang.invalidSoft)

def about():
    os.system("cls")
    print(lang.aboutMsg1)
    print(lang.aboutAuthor)
    print(lang.githubPage)
    print(lang.KMSinfo)

if select == "1":
    softSelect()
if select == "2":
    about()

#Windows Activation Confirmation
def WinActConfirm():
    os.system("cls")
    print(":=================Windows=================:\n ")
    print(lang.WinverAndNum + "{0} {1}\n ".format(sysNum,sysEdition))
    confirm = input(lang.startActInput)
    if confirm == "Д":
        StartWinAct()
    elif confirm == "д":
        StartWinAct()
    elif confirm == "Н":
        softSelect()
    elif confirm == "н":
        softSelect()
    elif confirm == "Y":
        StartWinAct()
    elif confirm == "y":
        StartWinAct()
    elif confirm == "N":
        softSelect()
    elif confirm == "n":
        softSelect()
    else:
        softSelect()



def OfficeVerSelect():
    pass

def OfficeActConfirm():
    os.system("cls")
    print(":=================Office=================:")
    print(lang.OfficeVerAndEdition)
    officeActConfirm = input(lang.OfficeStartActInput)
    if officeActConfirm == "Y":
        OfficeVerSelect()
    elif officeActConfirm == "N":
        softSelect()
    elif officeActConfirm == "y":
        OfficeVerSelect()
    elif officeActConfirm == "n":
        softSelect()
    elif officeActConfirm == "Д":
        OfficeVerSelect()
    elif officeActConfirm == "Н":
        softSelect()
    elif officeActConfirm == "д":
        OfficeVerSelect()
    elif officeActConfirm == "н":
        softSelect()
    else:
        softSelect()

def StartWinAct():
    srv = KMSdata.kmsServer
    """
    Definition of version and edition of Windows and activation.
    """

    #Windows 10
    if sysName == "10Professional":
        act.windows(srv, Windows10.Pro)
    elif sysName == "10ProfessionalN":
        act.windows(srv, Windows10.ProN)
    elif sysName == "10Home":
        act.windows(srv, Windows10.Home)
    elif sysName == "10HomeN":
        act.windows(srv, Windows10.HomeN)
    elif sysName == "10HomeSigleLanguage":
        act.windows(srv, Windows10.HomeSL)
    elif sysName == "10HomeCountrySpecific":
        act.windows(srv, Windows10.HomeCS)
    elif sysName == "10Education":
        act.windows(srv, Windows10.Edu)
    elif sysName == "10EducationN":
        act.windows(srv, Windows10.EduN)
    elif sysName == "10Enterprise":
        act.windows(srv, Windows10.Enterprise)
    elif sysName == "10EnterpriseN":
        act.windows(srv, Windows10.EnterpriseN)
    
    #Windows 8.1
    elif sysName == "8.1Professional":
        act.windows(srv, Windows81.Pro)
    elif sysName == "8.1ProfessionalN":
        act.windows(srv, Windows81.ProN)
    elif sysName == "8.1ProfessionalWMC":
        act.windows(srv, Windows81.ProWMC)
    elif sysName == "8.1Core":
        act.windows(srv, Windows81.Core)
    elif sysName == "8.1CoreN":
        act.windows(srv, Windows81.CoreN)
    elif sysName == "8.1CoreSingleLanguage":
        act.windows(srv, Windows81.CoreSL)
    elif sysName == "8.1Enterprise":
        act.windows(srv, Windows81.Enterprise)
    elif sysName == "8.1EnterpriseN":
        act.windows(srv, Windows81.EnterpriseN)
    
    #Windows 8
    elif sysName == "8Professional":
        act.windows(srv, Windows8.Pro)
    elif sysName == "8ProfessionalN":
        act.windows(srv, Windows8.ProN)
    elif sysName == "8ProfessionalWMC":
        act.windows(srv, Windows8.ProWMC)
    elif sysName == "8Core":
        act.windows(srv, Windows8.Core)
    elif sysName == "8Core SingleLagnuage":
        act.windows(srv, Windows8.CoreSL)
    elif sysName == "8Enterprise":
        act.windows(srv, Windows8.Enterprise)
    elif sysName == "8EnterpriseN":
        act.windows(srv, Windows8.EnterpriseN)
    
    else:
        os.system("cls")
        print(Back.RED+lang.OSnotSupport+Back.BLACK)
    




input()
