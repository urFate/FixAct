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
    def __init__(self, server, key):
        pass
    


        

