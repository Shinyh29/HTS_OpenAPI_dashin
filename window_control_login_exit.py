import pywinauto
from pywinauto import findwindows
from pywinauto.application import  Application
import os

def close_window_titlenm(titlenm='CPSTART'):
    procs = findwindows.find_elements()

    for proc in procs:
        print(proc.name)
        if proc.name == titlenm:
            print('-----------------------')
            app = Application(backend="uia").connect(process=proc.process_id) # process 연결
            dig = app[proc.name]
            print(dig)
            #os.system('tas')
            print(dig.window(title = f'{titlenm}').close())
            #print(dig.print_control_identifiers())
            #dig.Edit.type_keys('pywinauto{ENTER}test')
            #klick = dig['확인Button'].click()
            #eturn klick
            print('-----------------------')
close_window_titlenm(titlenm='CPSTART')