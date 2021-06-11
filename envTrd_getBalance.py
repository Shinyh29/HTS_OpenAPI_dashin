import win32com.client

instCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil") # 주문관련도구
instCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311") # 계좌정보
cpCash = win32com.client.Dispatch("CpTrade.CpTdNew5331A") # 주문가능금액


# instCpTdUtil (파이썬알고리즘트레이딩) == cpTradeUtil ( at 파이썬증권데이터분석 )

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

import sys
sys.flags = 0

get_balance = get_current_cash()
print(f'계좌잔고(주식) : {format(get_balance,",")} 원')