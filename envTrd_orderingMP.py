# Model Portfolio, MP 로 주문넣기
# 매일,  AI Portf 에 대한
## Ticker, Na



import win32com.client

instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil") # 주문관련도구
instCpTd0311= None
instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311") # 계좌정보
cpCash= None
cpCash = win32com.client.Dispatch("CpTrade.CpTdNew5331A") # 주문가능금액
instCpTdUtil.TradeInit()

# instCpTdUtil (파이썬알고리즘트레이딩) == cpTradeUtil ( at 파이썬증권데이터분석 )

def get_current_cash():
    global accFlag
    global acc
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


import pandas as pd
pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 100)

def load_signals():

    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    file_path = "C:\\PycharmProjects\\RL_trader_custom_csv\\portf_signals.xlsm"
    wb = xl.Workbooks.Open(file_path)
    sheet = wb.Sheets("portf_signals_T")
    sheet.Activate()
    tables = wb.ActiveSheet.Range("B5:D465").Value

    df_signalsAI3 = pd.DataFrame(tables)
    df_signalsAI3.rename(columns={0:'KR_nm',
                                  1:'Ticker',
                                  2:'Signal'},inplace=True)
    df_signalsAI3 = df_signalsAI3.iloc[1::]
    df_signalsAI3.reset_index(drop=True,inplace=True)
    print(df_signalsAI3)
    xl.Save
    # wb.Close(savechanges=1)
    xl.quit()
    os.system("taskkill /f /im excel.exe")
    # time.sleep(2)
    # pywinauto.keyboard.send_keys("{n}")
    # 저장하시곘습니까? 알림이떠버림

    df_signalsAI3['Signal'] = pd.to_numeric(df_signalsAI3['Signal'])
    # signals 숫자만  Float 형식으로 바꾸기

    return df_signalsAI3


# instCpPrice = win32com.client.Dispatch("DsCbo1.StockMst") # 현재가



def get_rlTimePx(ticker):
    instCpPrice = win32com.client.Dispatch("DsCbo1.StockMst")
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

# 현재가 잘안구해짐,.  차라리.  인터넷 Req로 얻자
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


def order_each_stockBuy(df):
    each_Buys = 0
    for idx, row in df.iterrows():
        if row.Signal != 0:
            print(idx,"\n", row)

            lets_buy = row.Signal * get_balance * 0.01
            each_Buy = float(round(lets_buy,-2))
            print(f'종목별_주문금액 : {format(round(lets_buy,-2),",")} 원' )

            # get_rltimePrice( Ticker )
            # sys.flags = 0
            get_px = get_realtime_stock(aws_ticker=str(row.Ticker))
            print(f'종목별_현재호가 : {get_px} 원')


            each_Buys += each_Buy
            # 종목별 주문금액
    print(f'==========\n종목전체_주문금액 : {format(each_Buys, ",")} 원')

import time

def order_buy(ticker, amnt, px):
    acc = instCpTdUtil.AccountNumber[0]
    instCpTd0311.SetInputValue(0, "2") # 2: Buy
    instCpTd0311.SetInputValue(1, acc)  # 계좌번호
    instCpTd0311.SetInputValue(2, accFlag[0]) # 상품구분 , 주식상품중 첫번쨰
    instCpTd0311.SetInputValue(3, ticker) # 종목
    instCpTd0311.SetInputValue(4, amnt) # 주문수량
    instCpTd0311.SetInputValue(5, px) # 주문단가
    instCpTd0311.SetInputValue(7,"0") # 주문조건 구분코드, {0: 기본, 1: IOC, 2: FOK}
    instCpTd0311.SetInputValue(8, "01") # 주문호가 구분코드 - 01: 보통

    # 매수 주문 요청
    nRet = instCpTd0311.BlockRequest()

    if nRet == 4:
        time.sleep(15.1)
        nRet = instCpTd0311.BlockRequest()

    if (nRet != 0):
        print("주문요청 오류", nRet)
        # 0: 정상,  그 외 오류, 4: 주문요청제한 개수 초과
        # 만약 4를 리턴받은 경우는 15초동안 호출 제한을 초과한 경우로 잠시 후 다시 요청이 필요
        exit()


    rqStatus = instCpTd0311.GetDibStatus()
    errMsg = instCpTd0311.GetDibMsg1()
    if rqStatus != 0:
        print("주문 실패: ", rqStatus, errMsg)
        exit()

    print(f'Odered ! Buy ----------{ticker},{amnt}, : {px}')

print(f'my balance : {get_balance} ')
import os
df = load_signals()
print(f'my balance : {get_balance} ')
# order_each_stockBuy(df=df)
import math
each_Buys = 0

if get_balance == 0:
    get_balance = 100000000

print(f'my balance : {get_balance} ')

for idx, row in df.iterrows():
    if row.Signal != 0:
        print("\n",idx,"====================\n", row)

        lets_buy = row.Signal * get_balance * 0.01
        each_Buy = float(round(lets_buy,-2))
        print(f'종목별_주문금액 : {format(round(lets_buy,-2),",")} 원' )

        # get_rltimePrice( Ticker )
        # sys.flags = 0
        ticker = row.Ticker
        get_px = get_realtime_stock(aws_ticker= ticker)
        print(f'종목별_현재호가 : {get_px} 원')

        # 주문할 수량
        buy_order_qt = round(lets_buy / get_px, 0)
        print(f'종목별_주문수량 : {math.floor(buy_order_qt)} 개')

        # 매수 주문하기
        order_buy(ticker= ticker, amnt=math.floor(buy_order_qt), px=get_px  )


        each_Buys += each_Buy
        # 종목별 주문금액
print(f'==========\n종목전체_주문금액 : {format(each_Buys, ",")} 원')

