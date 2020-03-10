import platform
import os
import colorama
from colorama import Back
import activ as act
from strings import Windows10, Windows8, Windows81, KMSdata, Office2010, Office2013, Office2016, Office2019
import locale
import wget
import time
import zipfile
colorama.init()

#Auto translate
loc = locale.getdefaultlocale()[0]
getlang = loc[0:2]
if getlang == "ru":
    from strings import langRU as lang
else:
    from strings import langEN as lang


#Welcome message
print(":===================FixAct 1.1===================:")
print(lang.welcome)
print(lang.mainSelect)
select = input(lang.InputSelect)

#Variables
sysVer = platform.version()
sysNum = platform.win32_ver()[0]
sysEdition = platform.win32_edition()
sysName = sysNum + sysEdition
#===============================

def StartOfficeAct():
    srv = KMSdata.kmsServer
    #Getting path to ProgramFiles folder
    progFilesFolder = os.environ["PROGRAMFILES"]

    #Office paths
    office10 = progFilesFolder+r"\Microsoft Office\Office14\ospp.vbs"
    office13 = progFilesFolder+r"\Microsoft Office\Office15\ospp.vbs"
    office = progFilesFolder+r"\Microsoft Office\Office16\ospp.vbs"

    if os.path.exists(office10):
        try:
            os.mkdir("C:/libraries")
            os.mkdir("C:/libraries/office10")
            lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2010.zip?raw=true"
            out = "C:/libraries/"
            wget.download(lib13, out=out)
            print(lang.officeLibDownload)
            libs = zipfile.ZipFile('C:/libraries/office2010.zip')
            libs.extractall('C:/libraries/office10/')
            os.chdir("C:/libraries/office13/library/")
            os.system(r"for /f %x in ('dir /b *') do cscript %folder%\ospp.vbs /inslic:%x")
            time.sleep(1.500)
            act.office(srv, Office2010.key, "10")
        except OSError:
            lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2010.zip?raw=true"
            out = "C:/libraries/"
            libs = zipfile.ZipFile('C:/libraries/office2010.zip')
            libs.extractall('C:/libraries/office10/')
            wget.download(lib13, out=out)
            print(lang.officeLibDownload)
            time.sleep(1.500)
            act.office(srv, Office2010.key, "10")
    elif os.path.exists(office13):
        try:
            os.mkdir("C:/libraries")
            os.mkdir("C:/libraries/office13")
            lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2013.zip?raw=true"
            out = "C:/libraries/"
            wget.download(lib13, out=out)
            print(lang.officeLibDownload)
            libs = zipfile.ZipFile('C:/libraries/office2013.zip')
            libs.extractall('C:/libraries/office13/')
            os.chdir("C:/libraries/office13/library/")
            os.system(r"for /f %x in ('dir /b *') do cscript %folder%\ospp.vbs /inslic:%x")
            time.sleep(1.500)
            act.office(srv, Office2013.key, "13")
        except OSError:
            lib13 = "https://github.com/Arny4/FixAct/blob/master/OldOfficeVL_libs/office2013.zip?raw=true"
            out = "C:/libraries/"
            libs = zipfile.ZipFile('C:/libraries/office2013.zip')
            libs.extractall('C:/libraries/office13/')
            wget.download(lib13, out=out)
            print(lang.officeLibDownload)
            time.sleep(1.500)
            act.office(srv, Office2013.key, "13")

    elif os.path.exists(office):
        print(lang.officeVerSelect)
        print("1) Office 2016(365)\n2) Office 2019")
        select = input(lang.InputSelect)
        if select == "1":
            act.office(srv, Office2016.key, "16")
        elif select == "2":
            act.office(srv, Office2019.key, "19")
    else:
        print(Back.RED+lang.outdatedOffice+Back.BLACK)

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



def OfficeActConfirm():
   #Getting path to ProgramFiles folder
    progFilesFolder = os.environ["PROGRAMFILES"]

    #Office versions variables
    office10 = progFilesFolder+r"\Microsoft Office\Office14\ospp.vbs"
    office13 = progFilesFolder+r"\Microsoft Office\Office15\ospp.vbs"
    office = progFilesFolder+r"\Microsoft Office\Office16\ospp.vbs"
    os.system("cls")
    print(":=================Office=================:")
    if os.path.exists(office10):
        print(lang.OfficeVerAndEdition+"Office 2010")
    elif os.path.exists(office13):
        print(lang.OfficeVerAndEdition+"Office 2013")
    elif os.path.exists(office):
        print(lang.OfficeVerAndEdition+"Office 2016 or higher")
    else:
        print("You dont have Office on your PC :c")
        time.sleep(3)
        softSelect()
    officeActConfirm = input(lang.OfficeStartActInput)
    if officeActConfirm == "Y":
        StartOfficeAct()
    elif officeActConfirm == "N":
        softSelect()
    elif officeActConfirm == "y":
        StartOfficeAct()
    elif officeActConfirm == "n":
        softSelect()
    elif officeActConfirm == "Д":
        StartOfficeAct()
    elif officeActConfirm == "Н":
        softSelect()
    elif officeActConfirm == "д":
        StartOfficeAct()
    elif officeActConfirm == "н":
        softSelect()
    else:
        softSelect()

#Activated Product Selection
def softSelect():
    os.system("cls")
    print(lang.softChoose)
    print("1) Windows\n2) Office\n ")
    softSelected = input(lang.softSelect)
    if softSelected == "1":
        WinActConfirm()
    if softSelected == "2":
        OfficeActConfirm()
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








input()
