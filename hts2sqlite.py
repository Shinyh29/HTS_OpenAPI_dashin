## 이유모르겠음. 에러발생

#!/usr/bin/env python
# coding: utf-8

# In[1]:


'''
1. 백테스트를 위한 데이터베이스 구축
- 데이터 조회
'''


# 패키지 호출 
import pandas as pd
import win32com.client
import sqlite3


# In[2]:


# 대신증권 연결여부 체크

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")


def check_connection():
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음.")
        exit()
        
    else:
        print(bConnect)
        print("연결완료")
        
    return True

check_connection()


# In[5]:


# 주가 불러오기 - 갯수 기준
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
 
objStockChart.SetInputValue(0, 'A005930')   #종목 코드 - 삼성전자

objStockChart.SetInputValue(1, ord('2')) # 개수로 조회
objStockChart.SetInputValue(4, 10) # 최근 100일 치

objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9]) #날짜,시가,고가,저가,종가,거래량
objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
objStockChart.BlockRequest()
 
len = objStockChart.GetHeaderValue(3)
 
print("날짜", "시가", "고가", "저가", "종가", "거래량")
print("==============================================")
 
for i in range(len):
    day = objStockChart.GetDataValue(0, i)
    open = objStockChart.GetDataValue(1, i)
    high = objStockChart.GetDataValue(2, i)
    low = objStockChart.GetDataValue(3, i)
    close = objStockChart.GetDataValue(4, i)
    vol = objStockChart.GetDataValue(5, i)
    print (day, open, high, low, close, vol)


# In[9]:


# 주가 불러오기 - 날짜 기준
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
 
objStockChart.SetInputValue(0, 'A005930') # 종목 코드 - 삼성전자
objStockChart.SetInputValue(1, ord('1')) # 날짜로 조회

objStockChart.SetInputValue(2, 0) # 종료 날짜, 0을 넣으면 가장 최근 날짜로 불러옴.
objStockChart.SetInputValue(3, 20200401) # 시작 날짜, 3월 1일로  설정하였음.

objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9]) # 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
objStockChart.BlockRequest()
 
len = objStockChart.GetHeaderValue(3)
 
print("날짜", "시가", "고가", "저가", "종가", "거래량", "거래대금")
print("==============================================")
 
for i in range(len):
    day = objStockChart.GetDataValue(0, i)
    open = objStockChart.GetDataValue(1, i)
    high = objStockChart.GetDataValue(2, i)
    low = objStockChart.GetDataValue(3, i)
    close = objStockChart.GetDataValue(4, i)
    vol = objStockChart.GetDataValue(5, i)
    vol_mon = objStockChart.GetDataValue(6, i)
    
    # 데이터 확인해보기
    print(day, open, high, low, close, vol, vol_mon)


# In[11]:


'''
- DataFrame로 저장하기
'''

# 1. 주가 불러오기 - 날짜 기준
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
 
objStockChart.SetInputValue(0, 'A005930') # 종목 코드 - 삼성전자
objStockChart.SetInputValue(1, ord('1')) # 날짜로 조회

objStockChart.SetInputValue(2, 0) # 종료 날짜, 0을 넣으면 가장 최근 날짜로 불러옴.
objStockChart.SetInputValue(3, 20200401) # 시작 날짜, 3월 1일로  설정하였음.

objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9]) # 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
objStockChart.SetInputValue(6, ord('D')) # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1')) # 수정주가 사용
objStockChart.BlockRequest()
 
len = objStockChart.GetHeaderValue(3)

# 2. 리스트 기반으로 저장
value_list = []

for i in range(len):
    day = objStockChart.GetDataValue(0, i)
    open = objStockChart.GetDataValue(1, i)
    high = objStockChart.GetDataValue(2, i)
    low = objStockChart.GetDataValue(3, i)
    close = objStockChart.GetDataValue(4, i)
    vol = objStockChart.GetDataValue(5, i)
    vol_mon = objStockChart.GetDataValue(6, i)
    
    # 데이터 확인해보기
    value_list.append([day, open, high, low, close, vol, vol_mon])

# 3, DataFrame로 변환
price_df = pd.DataFrame(value_list, columns = ['day','open', 'high', 'low', 'close', 'vol', 'vol_money'])

# 4. 데이터 확인
print(price_df)


# In[12]:


'''
방금 만든 데이터를 sqlite3 데이터베이스에 저장하기
'''

con = sqlite3.connect("price.db")
price_df.to_sql('A005930', con, if_exists='replace')


# In[ ]:




