# trader.py

import ctypes
import win32com.client

cpStatus = win32com.client.Dispatch('CpUtil.CpCybos') #시스템 상태정보
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil') # 주문관련 도구

# CREON Plus system Function
def check_creon_system():
    # 관리자권한 실행 check
    if not ctypes.windll.shell32.IsUserAnAdmin():
        print(f'check system () : admin -> Failed')
        return False

    if (cpStatus.IsConnect ==0):
        print(f'check creon system () : con2 server  -> Failed')
        return False

check_creon_system()