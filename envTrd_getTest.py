


import win32com.client
instCpPrice = win32com.client.Dispatch("DsCbo1.StockMst")

import time
import pythoncom
import win32event
StopEvent = win32event.CreateEvent(None, 0, 0, None)

def MessagePump(timeout):
    waitables = [StopEvent]
    while 1:
        rc = win32event.MsgWaitForMultipleObjects(
            waitables,
            0,  # Wait for all = false, so it waits for anyone
            timeout,  # (or win32event.INFINITE)
            win32event.QS_ALLEVENTS)  # Accepts all input

        if rc == win32event.WAIT_OBJECT_0:
            # Our first event listed, the StopEvent, was triggered, so we must exit
            print('stop event')
            break

        elif rc == win32event.WAIT_OBJECT_0 + len(waitables):
            # A windows message is waiting - take care of it. (Don't ask me
            # why a WAIT_OBJECT_MSG isn't defined < WAIT_OBJECT_0...!).
            # This message-serving MUST be done for COM, DDE, and other
            # Windowsy things to work properly!
            # print('pump')
            if pythoncom.PumpWaitingMessages():
                break  # we received a wm_quit message
        elif rc == win32event.WAIT_TIMEOUT:
            print('timeout')
            return
            pass
        else:
            print('exception')
            raise RuntimeError("unexpected win32wait return value")

def get_rlTimePx(ticker):

    instCpPrice.SetInputValue(0,ticker)
    instCpPrice.BlockRequest()
    # instCpPrice.Request()
    # instCpPrice.GetDibStatus()
    # instCpPrice.GetDibMsg1()
    # MessagePump(10000)
    a = instCpPrice.GetHeaderValue(0)
    b = instCpPrice.GetHeaderValue(1)
    c = instCpPrice.GetHeaderValue(11)

    print(a,b,c)
    return c

def get_rlTimePx_req(ticker):

    instCpPrice.SetInputValue(0,ticker)
    # instCpPrice.BlockRequest()
    instCpPrice.Request()
    # instCpPrice.GetDibStatus()
    # instCpPrice.GetDibMsg1()
    # MessagePump(10000)
    a = instCpPrice.GetHeaderValue(0)
    b = instCpPrice.GetHeaderValue(1)
    c = instCpPrice.GetHeaderValue(11)

    print(a,b,c)
    return c

import urllib3
urllib3req = urllib3.PoolManager()
import sys
import io
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding="utf-8")
sys.stdout = io.TextIOWrapper(sys.stderr.detach(), encoding="utf-8")
encoding = 'euc-kr'

import json

def get_realtime_stock(aws_ticker):
    aws_ticker = aws_ticker.replace("A","")
    url = f'https://polling.finance.naver.com/api/realtime?_callback=window.__jindo_callback._2805&query=SERVICE_ITEM%3A{aws_ticker}'
    user_agent = {"user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36"}
    headers = user_agent
    #html = requests.get(url, timeout=10).text
    html = urllib3req.request('GET',url).data
    # print(html)
    # html = requests.get(url, headers= headers, timeout=10).text
    #html = http.request('get',url,headers=headers)
    temp =str(html,encoding).split('"datas":')[1].split(']')[0].split('[')[1]
    temp = json.loads(temp)
    if - temp['sv'] + temp['nv'] > 0:
        idx = 1
    else:
        idx = -1
    temp['cr'] = idx * temp['cr']
    output = {'cd':temp['cd'],'nm':temp['nm'], 'pct':temp['cr'],'px':temp['nv'] }
    return output['px']


instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil") # 주문관련도구
instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311") # 계좌정보
cpCash = win32com.client.Dispatch("CpTrade.CpTdNew5331A") # 주문가능금액

def get_current_cash():
    instCpTdUtil.TradeInit()
    acc = instCpTdUtil.AccountNumber[0] # 계좌번호
    accFlag = instCpTdUtil.GoodsList(acc,1)
    # {-1:전체, 1: 주식, 2: 선물옵션 }
    cpCash.SetInputValue(0,acc)
    cpCash.SetInputValue(1, accFlag[0])
    cpCash.BlockRequest()

    print(f'계좌번호 :{acc}')
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문가능금액

get_balance = get_current_cash()
print(f'계좌잔고(주식) : {format(get_balance,",")} 원')

px1 =get_realtime_stock(aws_ticker="A005930")
px2 =get_realtime_stock(aws_ticker="A011200")
px3 =get_realtime_stock(aws_ticker="A000060")

print(f'=========={px1}, {px2}, {px3}')