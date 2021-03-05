import win32com.client
import pandas as pd
pd.set_option('display.max_columns', 100)

ticker = "A005930"

def ReqeustData(obj):
    # 데이터 요청
    obj.BlockRequest()

    # 통신 결과 확인
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False

    # 일자별 정보 데이터 처리
    count = obj.GetHeaderValue(1)  # 데이터 개수
    for i in range(count):
        date = obj.GetDataValue(0, i)  # 일자
        open = obj.GetDataValue(1, i)  # 시가
        high = obj.GetDataValue(2, i)  # 고가
        low = obj.GetDataValue(3, i)  # 저가
        close = obj.GetDataValue(4, i)  # 종가
        diff = obj.GetDataValue(5, i)  # 종가
        vol = obj.GetDataValue(6, i)  # 종가
        #print(date, open, high, low, close, diff, vol)

    return True


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()


## 방법2 일간조회

# 주가 불러오기 - 날짜 기준
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
# 관심종목 객체 구하기
#objStockChart = win32com.client.Dispatch('CpSysDib.CpSvr7254')


objStockChart.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자
objStockChart.SetInputValue(1, ord('1'))  # 날짜로 조회

objStockChart.SetInputValue(2, 0)  # 종료 날짜, 0을 넣으면 가장 최근 날짜로 불러옴.
objStockChart.SetInputValue(3, 20150101)  # 시작 날짜, 3월 1일로  설정하였음.

objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9, 12,13,17,20,21])  # 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
"""
0~9 : # 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
{12:상장주식수, 13: 시가총액, 17 : 외국인현재보유비율, 20: 기관순매수, 21: 기관누적순매수  ? }
"""


objStockChart.SetInputValue(6, ord('D'))  # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
objStockChart.BlockRequest()

len = objStockChart.GetHeaderValue(3)

print("날짜", "시가", "고가", "저가", "종가", "거래량", "거래대금")
print("12:상장주식수, 13: 시가총액, 17 : 외국인현재보유비율, 20: 기관순매수, 21: 기관누적순매수 ")
print("==============================================")

value_list = []

for i in range(len):
    day = objStockChart.GetDataValue(0, i)
    open = objStockChart.GetDataValue(1, i)
    high = objStockChart.GetDataValue(2, i)
    low = objStockChart.GetDataValue(3, i)
    close = objStockChart.GetDataValue(4, i)
    vol = objStockChart.GetDataValue(5, i)
    vol_mon = objStockChart.GetDataValue(6, i)

    ### 이후 추가데이터
    data7 = objStockChart.GetDataValue(7, i)
    data8 = objStockChart.GetDataValue(8, i)
    data9 = objStockChart.GetDataValue(9, i)
    data10 = objStockChart.GetDataValue(10, i)


    # 데이터 확인해보기
    print(day, open, high, low, close, vol, vol_mon)
    value_list.append([day, open, high, low, close, vol, vol_mon, \
                       data7, data8, data9, data10])



price_df = pd.DataFrame(value_list)
price_df.columns = ["Date", "Open", "High", "Low", "Close", "Volume", "trs_volume", \
                    'num_listed', 'cap', 'ratio_foreigner', 'instit_net_buy']
#print(price_df[['Date','num_listed', 'cap', 'ratio_foreigner', 'instit_net_buy','instit_net_budycum']])
#print(price_df)
ratio_trs_volume = price_df['trs_volume'] / price_df['cap']
ratio_instit_netbuy = price_df['instit_net_buy'] / price_df['cap']
ratio_foreign_netbuy = price_df['ratio_foreigner']

df_insert2sql = pd.DataFrame()
df_insert2sql['ratio_trs_volume'] = ratio_trs_volume
df_insert2sql['ratio_instit_netbuy'] = ratio_instit_netbuy * 10 ** 8 # 1억 " 1* 10^ 8
df_insert2sql['ratio_foreign_netbuy'] = ratio_foreign_netbuy
print(df_insert2sql)


# In[11]:


'''
- DataFrame로 저장하기
'''

# 1. 주가 불러오기 - 날짜 기준

## 전체종목 넣기
# import tickers_list
# print(tickers_list.dfs['Ticker'])
#
# for unit in tickers_list.dfs['Ticker']:
#     def insert2tb(unit):
#         objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
#         objStockChart.SetInputValue(0, unit)  # 종목 코드 - 삼성전자
#         objStockChart.SetInputValue(1, ord('1'))  # 날짜로 조회
#
#         objStockChart.SetInputValue(2, 0)  # 종료 날짜, 0을 넣으면 가장 최근 날짜로 불러옴.
#         objStockChart.SetInputValue(3, 20210219)  # 시작 날짜, 3월 1일로  설정하였음.
#
#         objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9])  # 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
#         objStockChart.SetInputValue(6, ord('D'))  # '차트 주가 - 일간 차트 요청
#         objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
#         objStockChart.BlockRequest()
#
#         len = objStockChart.GetHeaderValue(3)
#
#         # 2. 리스트 기반으로 저장
#         value_list = []
#
#         for i in range(len):
#             day = objStockChart.GetDataValue(0, i)
#             open = objStockChart.GetDataValue(1, i)
#             high = objStockChart.GetDataValue(2, i)
#             low = objStockChart.GetDataValue(3, i)
#             close = objStockChart.GetDataValue(4, i)
#             vol = objStockChart.GetDataValue(5, i)
#             vol_mon = objStockChart.GetDataValue(6, i)
#
#
#             # 데이터 확인해보기
#             value_list.append([day, open, high, low, close, vol, vol_mon])
#             #value_list.append([unit, day, close])
#             ## unit = ticker
#             price_df = pd.DataFrame(value_list)
#             print(price_df.head())
#


        # # 3, DataFrame로 변환
        # price_df = pd.DataFrame(value_list, columns=[ 'Ticker','Date', 'Value'])
        # #
        # # # 4. 데이터 확인
        # print(f'-------{unit}-----')
        # print(price_df.head())

        # from sqlalchemy import create_engine
        #
        # item_tb = 'tickers'
        # pw ='0000'
        # ip_public = '13.209.4.191'
        # port = '3306'
        # db_name = 'ssiaat_shin'
        # engine = create_engine("mysql+pymysql://root:" + pw + f"@{ip_public}:{port}/{db_name}?charset=utf8",
        #                    encoding='utf-8')
        #
        #
        #
        #
        #
        # try:
        #     price_df.to_sql(name='px_close', con=engine, if_exists='append', index=False)
        # except Exception as e:
        #     print(f'{e}_____Failed to bulkdf 2 EC2 insert')


    #insert2tb(unit)

