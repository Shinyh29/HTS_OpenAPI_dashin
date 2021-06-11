import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
import pandas as pd

import date_set
today = date_set.today(0)
################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0

    def Request(self, codes, dataInfo):
        # 0: 종목코드 4: 현재가 20: 상장주식수
        rqField = [0, 4, 20]  # 요청 필드

        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(2)

        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            cur = self.objRq.GetDataValue(1, i)  # 종가
            listedStock = self.objRq.GetDataValue(2, i)  # 상장주식수

            maketAmt = listedStock * cur
            if g_objCodeMgr.IsBigListingStock(code):
                maketAmt *= 1000
            #            print(code, maketAmt)

            # key(종목코드) = tuple(상장주식수, 시가총액)
            dataInfo[code] = (listedStock, maketAmt)

        return True


class CMarketTotal():
    def __init__(self):
        self.dataInfo = {}

    def GetAllMarketTotal(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        allcodelist = codeList + codeList2
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d' % (len(allcodelist), len(codeList), len(codeList2)))

        objMarket = CpMarketEye()
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo)
                rqCodeList = []
                continue
        # end of for

        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo)

    def PrintMarketTotal(self):
        global df
        # 시가총액 순으로 소팅
        data2 = sorted(self.dataInfo.items(), key=lambda x: x[1][1], reverse=True)
        df = pd.DataFrame(data2)

        print('전종목 시가총액 순 조회 (%d 종목)' % (len(data2)))

        names = []
        num_stocks = []
        caps =[]
        for item in data2:
            name = g_objCodeMgr.CodeToName(item[0])
            listed = item[1][0]
            markettot = item[1][1]
            print('%s 상장주식수: %s, 시가총액 %s' % (name, format(listed, ','), format(markettot, ',')))
            names.append(name)
            num_stocks.append(listed)
            caps.append(markettot)

        df = df.rename(columns= {0:'Ticker', 1:'temp'})
        df['name'] = names
        df['num_stock'] = num_stocks
        df['cap'] = caps
        df['Date'] = today
        df = df.drop(columns='temp')
        df = df[['Ticker','Date','cap','num_stock','name']]
        print(df)





if __name__ == "__main__":
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()

    import pandas as pd
    import pymysql

    import time
    from sqlalchemy import create_engine

    pw = '0000'
    ip_public = '13.209.4.191'
    port = '3306'
    db_name = 'ssiaat_shin'

    engine = create_engine("mysql+pymysql://root:" + pw + f"@{ip_public}:{port}/{db_name}?charset=utf8",
                           encoding='utf-8')

    item_tb = 'listed_cap'


    def TableCreater(item_tb):
        print(f''' make Table''')
        conn = pymysql.connect(host=ip_public, port=3306, user='root', password=pw, db=db_name,
                               charset='utf8')
        with conn.cursor() as curs:
            sql = f"""
                CREATE TABLE {item_tb} (
                Ticker VARCHAR(30) NOT NULL,
                Date Date NOT NULL,
                cap VARCHAR(64) NOT NULL,
                num_stock VARCHAR(64) NOT NULL,
                name VARCHAR(64) NOT NULL,
                );
                """
            # PRIMARY KEY(Ticker, Date)
            # Date DATE NOT NULL,
            curs.execute(sql)
            print(f'{sql}')
            conn.commit()
            codes = dict()
            conn.close()
    try:
        TableCreater(item_tb)
    except Exception as e:
        print(f'{e}')
        # 기존테이블 전체 삭제 후 업데이트
        conn1 = pymysql.connect(host='13.209.4.191', port=3306, user='root', password='0000', db='ssiaat_shin',
                                charset='utf8')
        query0 = 'USE ssiaat_shin;'
        query1 = f'TRUNCATE TABLE {item_tb};'
        curs = conn1.cursor()
        curs.execute(query0)
        curs.execute(query1)

    try:
        df.to_sql(name=item_tb, con=engine, if_exists='append', index=False)
        #df.to_excel('tickers_all.xlsx')
    except Exception as e:
        print(f'{e}_____Failed to bulkdf 2 EC2 insert')
