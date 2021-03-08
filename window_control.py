import pywinauto
from pywinauto import findwindows
from pywinauto.application import  Application


def close_window_titlenm(titlenm='CPSYSDIB'):
    procs = findwindows.find_elements()

    for proc in procs:
        if proc.name == titlenm:
            app = Application(backend="uia").connect(process=proc.process_id) # process 연결
            dig = app[proc.name]
            #print(dig.print_control_identifiers())
            #dig.Edit.type_keys('pywinauto{ENTER}test')
            klick = dig['확인Button'].click()
            return klick

close_window_titlenm(titlenm='CPSYSDIB')