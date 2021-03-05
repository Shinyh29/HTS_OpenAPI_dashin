# https://wikidocs.net/3299
import sys
from PyQt5.QtWidgets import *
import win32com.client
from pandas import Series, DataFrame
import pandas as pd
pd.set_option('display.max_columns', 100)
import time

# ### CpSvr7254  투자주체별현황을 일별/기간별, 순매수/매매비중을 일자별로 확인할수 있습니다
def subCpSvr7254(m_code, m_FromDate, m_ToDate):
   # ## 대신 API 세팅
    cpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7254")
    cpSvr7254.SetInputValue(0, m_code)       # 종목코드
    cpSvr7254.SetInputValue(1, '0')          # 기간선택 0:기간선택, 1:1개월, ... , 4:6개월
    cpSvr7254.SetInputValue(2, m_FromDate)  # 시작일자
    cpSvr7254.SetInputValue(3, m_ToDate)    # 끝일자
    cpSvr7254.SetInputValue(4, '0')         # 0:순매수 1:비중
    cpSvr7254.SetInputValue(5, '0')         # 투자자
    cpSvr7254.BlockRequest()

    numData=cpSvr7254.GetHeaderValue(1)
    # print(numData)
    data=[]
    for ixRow in range(numData):
        tempData=[]
        for ixCol in range(14):
            tempData.append(cpSvr7254.GetDataValue(ixCol, ixRow))
        data.append(tempData)

    # 연속 수행
    while cpSvr7254.Continue:
        cpSvr7254.BlockRequest()
        numData = cpSvr7254.GetHeaderValue(1)
        # print(numData)
        for ixRow in range(numData):
            tempData=[]
            for ixCol in range(14):
                tempData.append(cpSvr7254.GetDataValue(ixCol, ixRow))
            data.append(tempData)
        time.sleep(0.1)

    return data

# from subDS import subCpSvr7254
from pandas import DataFrame

import date_set


if __name__ == "__main__":
    # max rows request : 250 일 .  :: 1y ?
    #  from , toDate 를 수정해야함
    code='A005930'        # 삼성전자 코드  #
    fromDate = date_set.today(-300) # 요청 시작 날짜
    toDate = date_set.today(-1)   # 요청 마지막 날짜
    print(f'from {fromDate} ~ to {toDate}')
    # ### 자료가져오기
    data=subCpSvr7254(code, fromDate, toDate)
    print(data)
    df=DataFrame(data,  columns=['Date', '개인', '외국인', '기관계', '금융투자', '보험', '투신', '은행', '기타금융', '연기금', '기타법인', '기타외인', '사모펀드', '국가지자체'])
    print(df)
    print(df.Date.iloc[-1])  # 마지막날을  요청 첫번로 다음으로 넣어야함