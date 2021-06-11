# Model Portfolio, MP 로 주문넣기
# 매일,  AI Portf 에 대한
## Ticker, Na



import win32com.client

instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil") # 주문관련도구
instCpTd0311= None
instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311") # 계좌정보
cpCash= None
cpCash = win32com.client.Dispatch("CpTrade.CpTdNew5331A") # 주문가능금액
instCPTd6033 = None
instCPTd6033 = win32com.client.Dispatch("CpTrade.CpTd6033") # 매도할떄 ?

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


def order_fit(ticker, amnt, px):
    acc = instCpTdUtil.AccountNumber[0]
    if amnt > 0:
        longshort = 2 # Buy
        longshort_text = "Buy"
    elif amnt < 0:
        longshort = 1 # Sell
        amnt = abs(amnt)
        longshort_text = "Sell"

    instCpTd0311.SetInputValue(0, str(longshort)) # 1:Sell,  2: Buy
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

    print(f'Odered ! {longshort_text} ----------{ticker},{amnt}, : {px}')


def check_ihave():
    global items
    global df_ihave
    instCPTd6033.SetInputValue(0, acc) # 계좌번호
    instCPTd6033.SetInputValue(1, accFlag[0])   # 상품구분 , 주식상품중 첫번쨰
    instCPTd6033.SetInputValue(2, 50)  # 요청건수 최대 50

    df_ihave = pd.DataFrame()
    Tickers = []
    KRnames = []
    ihaves = []
    canSells = []
    avgs = []

    # 아래 반복  ( 50 이상으로 넘어가야함 )
    while True:
        ret = instCPTd6033.BlockRequest()
        if ret ==4:
            time.sleep(15)
            ret = instCPTd6033.BlockRequest()
        if ret !=4:
            print("보유수량 조회 고고 ")
        # 통신 및 통신 에러 처리
        rqStatus = instCPTd6033.GetDibStatus()
        rqRet = instCPTd6033.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
        cnt = instCPTd6033.GetHeaderValue(7)  # 보유 종류수
        print(f'========== 보유 종류수 : cnt : {cnt} 개')

        items = []
        for i in range(cnt):
            item = {}
            code = instCPTd6033.GetDataValue(12, i) # Ticker
            item['Ticker'] = code
            item['KRname'] = instCPTd6033.GetDataValue(0, i) # KRName
            item['ihave'] = instCPTd6033.GetDataValue(7, i) # 체결잔고수량
            item['canSell'] = instCPTd6033.GetDataValue(15, i) # 매도가능수량
            item['avg'] = instCPTd6033.GetDataValue(17,i) # 체결장부단가

            print(i,",", item)
            items.append(item)
            Tickers.append(item['Ticker'])
            KRnames.append(item['KRname'])
            ihaves.append(item['ihave'])
            canSells.append(item['canSell'])
            avgs.append(item['avg'])

            df_unit = pd.DataFrame()
            df_unit['Ticker'] = Tickers
            df_unit['KRname'] = KRnames
            df_unit['ihave'] = ihaves
            df_unit['canSell'] = canSells
            df_unit['avg'] = avgs

        df_ihave = pd.concat([df_ihave, df_unit], axis=0, ignore_index=False)

        # if len(items) > 51:
        #     break
        print(f" instCPTd6033.Continue : {instCPTd6033.Continue}")
        if (instCPTd6033.Continue == False):
            print(f" instCPTd6033.Continue : {instCPTd6033.Continue}")
            break

    return df_ihave.reset_index(drop=True)



g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')

def requestJango( caller):
    while True:
        ret = instCpTd0311.BlockRequest()
        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('연속조회 제한 오류, 남은 시간', remainTime)
            return False
        # 통신 및 통신 에러 처리
        rqStatus = instCpTd0311.GetDibStatus()
        rqRet = instCpTd0311.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = instCpTd0311.GetHeaderValue(7)
        print(cnt)

        for i in range(cnt):
            item = {}
            code = instCpTd0311.GetDataValue(12, i)  # 종목코드
            item['종목코드'] = code
            item['종목명'] = instCpTd0311.GetDataValue(0, i)  # 종목명
            item['현금신용'] = self.objRq.GetDataValue(1, i)  # 신용구분
            print(code, '현금신용', item['현금신용'])
            item['대출일'] = instCpTd0311.GetDataValue(2, i)  # 대출일
            item['잔고수량'] = instCpTd0311.GetDataValue(7, i)  # 체결잔고수량
            item['매도가능'] = instCpTd0311.GetDataValue(15, i)
            item['장부가'] = instCpTd0311.GetDataValue(17, i)  # 체결장부단가
            # 매입금액 = 장부가 * 잔고수량
            item['매입금액'] = item['장부가'] * item['잔고수량']

            # 잔고 추가
            caller.jangoData[code] = item

            if len(caller.jangoData) >= 200:  # 최대 200 종목만,
                break

        if len(caller.jangoData) >= 200:
            break
        if (self.objRq.Continue == False):
            break
    return True



def order_sell(ticker, amnt, px):
    acc = instCpTdUtil.AccountNumber[0]
    instCpTd0311.SetInputValue(0, "1") # 1:Sell, 2: Buy
    instCpTd0311.SetInputValue(1, acc)  # 계좌번호
    instCpTd0311.SetInputValue(2, accFlag[0]) # 상품구분 , 주식상품중 첫번쨰
    instCpTd0311.SetInputValue(3, ticker) # 종목
    instCpTd0311.SetInputValue(4, amnt) # 주문수량
    instCpTd0311.SetInputValue(5, px) # 주문단가
    instCpTd0311.SetInputValue(7,"0") # 주문조건 구분코드, {0: 기본, 1: IOC, 2: FOK}
    instCpTd0311.SetInputValue(8, "01") # 주문호가 구분코드 - 01: 보통

    # 매수, 매도 주문 요청
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

print(f'my balance : {get_balance} ')
# order_each_stockBuy(df=df)
import math
each_Buys = 0

if get_balance == 0:
    get_balance = 90000000

# 강제로  Fit 시키기
get_balance = 120000000
print(f'my balance : {get_balance} ')

# 내가 가진 종목별 수량 체크
df_ihave = check_ihave()


print(f'=========== get df_ihave :\n{df_ihave}')
df = load_signals()
print(f'===========Check Point ')
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

        # 매수/매도 주문상태
        my_status = 1
        ## my_status =
        ## {0 : Only KRW, 1: Check Ihave,  All Sell ,  2: Ceck Ihave fit to MP}
        if my_status ==0:  # When Only iIhave KRW
            order_buy(ticker= ticker, amnt=math.floor(buy_order_qt), px=get_px  )

        # 매도 주문하기
        if my_status ==1:  #Check Ihave
            print(f'=========== Check [종목별 ] 보유수량 확인 ')
            try:
                each_ihave = df_ihave[df_ihave.Ticker == ticker].ihave.values[0]
                # print(each_ihave)
            except Exception as e:
                # print(e)
                each_ihave = 0
            if each_ihave == None:
                each_ihave = 0
            print(f'종목별_보유수량 : {each_ihave} 개')

            # 종목별수량 -  현재보유수량 = 주문해야할 수량
            # if > 0 : 매수
            # if < 0 : 매도
            order_delta =  math.floor(buy_order_qt) - each_ihave
            print(f'종목별_주문할수량 : {order_delta} 개')
            if order_delta != 0:
                # 주문할 수량이  0 이 아닐때인 종목만 주문하기
                order_fit(ticker=ticker,amnt=order_delta,px=get_px)

        each_Buys += each_Buy
        # 종목별 주문금액
print(f'==========\n종목전체_주문금액 : {format(each_Buys, ",")} 원')

