## 종목별  거래대금  파이썬] 매매입체분석(투자주체별현황) 예제

import sys
from PyQt5.QtWidgets import *
import win32com.client
from pandas import Series, DataFrame
import pandas as pd
pd.set_option('display.max_columns', 100)
import locale
import os
import time

locale.setlocale(locale.LC_ALL, '')
# cp object
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

gExcelFile = '7254.xlsx'


class CpRp7354:
    def Request(self, code):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print('PLUS가 정상적으로 연결되지 않음. ')
            return False

        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch('CpSysDib.CpSvr7254')
        objRq.SetInputValue(0, code)
        objRq.SetInputValue(1, 6)  # 일자별
        objRq.SetInputValue(4, ord('0'))  # '0' 순매수 '1' 매매비중
        objRq.SetInputValue(5, 0)  # '전체
        objRq.SetInputValue(6, ord('1'))  # '1' 순매수량 '2' 추정금액(백만)

        sumcnt = 0
        data7254 = None
        data7254 = pd.DataFrame(columns=('date', 'close', '개인', '외국인', '기관계',
                                                '금융투자', '보험', '투신', '은행', '기타금융', '연기금', '국가,지자체',
                                                '기타법인', '기타외인'))

        while True:
            remainCount = g_objCpStatus.GetLimitRemainCount(1)  # 1 시세 제한
            if remainCount <= 0:
                print('시세 연속 조회 제한 회피를 위해 sleep', g_objCpStatus.LimitRequestRemainTime)
                time.sleep(g_objCpStatus.LimitRequestRemainTime / 1000)

            objRq.BlockRequest()

            # 현재가 통신 및 통신 에러 처리
            rqStatus = objRq.GetDibStatus()
            print('통신상태', rqStatus, objRq.GetDibMsg1())
            if rqStatus != 0:
                return False

            cnt = objRq.GetHeaderValue(1)
            sumcnt += cnt

            for i in range(cnt):
                item = {}
                item['date'] = objRq.GetDataValue(0, i)
                item['close'] = objRq.GetDataValue(14, i)
                item['개인'] = objRq.GetDataValue(1, i)
                item['외국인'] = objRq.GetDataValue(2, i)
                item['기관계'] = objRq.GetDataValue(3, i)
                item['금융투자'] = objRq.GetDataValue(4, i)
                item['보험'] = objRq.GetDataValue(5, i)
                item['투신'] = objRq.GetDataValue(6, i)
                item['은행'] = objRq.GetDataValue(7, i)
                item['기타금융'] = objRq.GetDataValue(8, i)
                item['연기금'] = objRq.GetDataValue(9, i)
                item['국가,지자체'] = objRq.GetDataValue(13, i)
                item['기타법인'] = objRq.GetDataValue(10, i)
                item['기타외인'] = objRq.GetDataValue(11, i)

                data7254.loc[len(data7254)] = item

            # 1000 개 정도만 처리
            if sumcnt > 1000:
                break;
            # 연속 처리
            if objRq.Continue != True:
                break

        data7254 = data7254.set_index('date')
        # 인덱스 이름 제거
        data7254.index.name = None
        print(data7254)
        return True



data7254 = DataFrame()

#obj7254 = CpRp7354()
#obj7254.Request('A005930', caller.data7254 )

print(CpRp7354.Request(self= None, code="A005930"))

