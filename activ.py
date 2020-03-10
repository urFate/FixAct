import os
import colorama
from colorama import Back
import locale
colorama.init()

#Auto translate
loc = locale.getdefaultlocale()[0]
getlang = loc[0:2]
if getlang == "ru":
    from strings import langRU as lang
else:
    from strings import langEN as lang

class windows:
    def __init__(self,server,key):
        os.system("cls")
        print(Back.CYAN+lang.actStarted+Back.BLACK) 
        os.system("cscript %windir%/system32/slmgr.vbs /ipk {0}".format(key)) #Изменяем ключ Windows
        os.system("cscript %windir%/system32/slmgr.vbs /skms {0}".format(server)) #Выставляем KMS сервер для активации
        os.system("cscript %windir%/system32/slmgr.vbs /ato") #Посылаем запрос на активацию
        input()

class office:
    def __init__(self, server, key, ver):
        #Getting path to ProgramFiles folder
        progFilesFolder = os.environ["PROGRAMFILES"]

        #Office paths
        office10 = progFilesFolder+r"\Microsoft Office\Office14\ospp.vbs"
        office13 = progFilesFolder+r"\Microsoft Office\Office15\ospp.vbs"
        office = progFilesFolder+r"\Microsoft Office\Office16\ospp.vbs"
        
        if ver == "10":
            os.system("cls")
            print(Back.CYAN+lang.actStarted+Back.BLACK) 
            os.system("cscript {0} /inpkey:{1}".format(office10, key))
            os.system("cscript {0} /sethst:{1}".format(office10, server))
            os.system("cscript {0} /setprt:1688".format(office10))
            os.system("cscript {0} /act".format(office10))
        elif ver == "13":
            os.system("cls")
            print(Back.CYAN+lang.actStarted+Back.BLACK) 
            os.system("cscript {0} /inpkey:{1}".format(office13, key))
            os.system("cscript {0} /sethst:{1}".format(office13, server))
            os.system("cscript {0} /setprt:1688".format(office13))
            os.system("cscript {0} /act".format(office13))
        elif ver == "16":
            os.system("cls")
            print(Back.CYAN+lang.actStarted+Back.BLACK) 
            os.system("cscript {0} /inpkey:{1}".format(office, key))
            os.system("cscript {0} /sethst:{1}".format(office, server))
            os.system("cscript {0} /setprt:1688".format(office))
            os.system("cscript {0} /act".format(office))
        elif ver == "19":
            os.system("cls")
            print(Back.CYAN+lang.actStarted+Back.BLACK) 
            os.system("cscript {0} /inpkey:{1}".format(office, key))
            os.system("cscript {0} /sethst:{1}".format(office, server))
            os.system("cscript {0} /setprt:1688".format(office))
            os.system("cscript {0} /act".format(office))



    


        

