import win32com.client
import pandas as pd
pd.set_option('display.max_columns', 100)

ticker = "A035420"


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

def get_chart_daily(code, start_date):
    objStockChart.SetInputValue(0, code)  # 종목 코드 - 삼성전자
    objStockChart.SetInputValue(1, ord('1'))  # 날짜로 조회
    ### 종료날짜 루프 필요 1665 에 1 request
    objStockChart.SetInputValue(2, 0)  # 종료 날짜, 0을 넣으면 가장 최근 날짜로 불러옴.
    objStockChart.SetInputValue(3, start_date) # 200000101)  # 시작 날짜, 3월 1일로  설정하였음.
    objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8, 9, 12,13,17,20,21])  # 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
    """
0: 날짜(ulong)
1:시간(long) - hhmm
2:시가(long or float)
3:고가(long or float)
4:저가(long or float)
5:종가(long or float)
6:전일대비(long or float) - 주) 대비부호(37)과 반드시 같이 요청해야 함
8:거래량(ulong or ulonglong) 주) 정밀도 만원 단위
9:거래대금(ulonglong)
10:누적체결매도수량(ulong or ulonglong) - 호가비교방식 누적체결매도수량
11:누적체결매수수량(ulong or ulonglong) - 호가비교방식 누적체결매수수량
 (주) 10, 11 필드는 분,틱 요청일 때만 제공
12:상장주식수(ulonglong)
13:시가총액(ulonglong)
14:외국인주문한도수량(ulong)
15:외국인주문가능수량(ulong)
16:외국인현보유수량(ulong)
17:외국인현보유비율(float)
18:수정주가일자(ulong) - YYYYMMDD
19:수정주가비율(float)
20:기관순매수(long)
21:기관누적순매수(long)
22:등락주선(long)
23:등락비율(float)
24:예탁금(ulonglong)
25:주식회전율(float)
26:거래성립률(float)
37:대비부호(char) - 수신값은 GetHeaderValue 8 대비부호와 동일
    """
    objStockChart.SetInputValue(6, ord('D'))  # '차트 주가 - 일간 차트 요청
    objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
    print(f'=========================Block Request.Start')
    df = pd.DataFrame()

    # ------------------------------


    objStockChart.BlockRequest()
    len = objStockChart.GetHeaderValue(3)

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
        #print(day, open, high, low, close, vol, vol_mon)
        value_list.append([day, open, high, low, close, vol, vol_mon, \
                           data7, data8, data9, data10])



    price_df = pd.DataFrame(value_list)
    price_df.columns = ["Date", "Open", "High", "Low", "Close", "Volume", "trs_volume", \
                        'num_listed', 'cap', 'rate_foreign', 'netbuy_instit']
        #print(price_df[['Date','num_listed', 'cap', 'ratio_foreigner', 'instit_net_buy','instit_net_budycum']])
    #print(price_df)

    df = price_df




    while objStockChart.Continue:
        objStockChart.BlockRequest()
        len = objStockChart.GetHeaderValue(3)
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
            #print(day, open, high, low, close, vol, vol_mon)
            value_list.append([day, open, high, low, close, vol, vol_mon, \
                           data7, data8, data9, data10])



        price_df = pd.DataFrame(value_list)
        price_df.columns = ["Date", "Open", "High", "Low", "Close", "Volume", "trs_volume", \
                        'num_listed', 'cap', 'rate_foreign', 'netbuy_instit']
        #print(price_df[['Date','num_listed', 'cap', 'ratio_foreigner', 'instit_net_buy','instit_net_budycum']])
        #print(price_df)
        df = pd.concat([df, price_df],axis=0, ignore_index=False)

    df['Date'] = pd.to_datetime(df['Date'].astype(str), format='%Y-%m-%d')
    df.set_index(['Date'], inplace=True)
    #print(f'-----------------total df : \n{df}')

    return df

# df = get_chart_daily(code=ticker, start_date="20000101")
# print(f'-----------------total df : \n \
# {df}')
# # start_date="20000101" : yyyymmdd
# print(df.keys().tolist())
