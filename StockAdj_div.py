import win32com.client
import pandas as pd
pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 30)

ticker = "A005930"


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()


## 방법2 일간조회



# 주가 불러오기 - 날짜 기준
objStockChart = win32com.client.Dispatch("CpSysDib.StockAdj")
# 관심종목 객체 구하기
#objStockChart = win32com.client.Dispatch('CpSysDib.CpSvr7254')

def get_chart_daily(code, start_date):
    objStockChart.SetInputValue(0, code)  # 종목 코드 - 삼성전자
    objStockChart.SetInputValue(1,'D')  # 날짜로 조회
    ### 종료날짜 루프 필요 1665 에 1 request
    """

    """
    print(f'=========================Block Request.Start')
    df = pd.DataFrame()

    # ------------------------------


    objStockChart.BlockRequest()
    len = objStockChart.GetHeaderValue(2)

    print("==============================================")
    value_list = []
    for i in range(len):
        day = objStockChart.GetDataValue(0, i)
        lock_info = objStockChart.GetDataValue(3, i)
        lock_before = objStockChart.GetDataValue(6, i)
        lock_after = objStockChart.GetDataValue(7, i)
        #print(high3)
        """
    3 - (string) 락구분코드
    00:해당사항없음(락이 발생안한 경우), 01:권리락,02:배당락,03:분배락,04:권배락,05:중간(분기)배당락,06:권리중간배당락,07:권리분기배당락,99:기타
    
    4 - (string) 액면가변경구분코드
    00:해당없음, 01:액면분할, 02:액면병합, 03:주식분할, 04:주식병합, 99:기타
    
    5 - (string) 재평가종목사유코드
    00:해당없음, 01:회사분할, 02:자본감소, 03:장기간정지, 04:초과분배,05:대규모배당, 06: 회사분할합병, 99:기타
    
    6 - (long) 변경전 기준가
    7 - (long) 변경후 기준가
        """

        # 데이터 확인해보기
        #print(day, open, high, low, close, vol, vol_mon)
        value_list.append([day, lock_info, lock_before, lock_after ])


    #print(value_list)
    price_df = pd.DataFrame(value_list)
    price_df.columns = ["Date", "lock_info", "lock_before", "lock_after"]
        #print(price_df[['Date','num_listed', 'cap', 'ratio_foreigner', 'instit_net_buy','instit_net_budycum']])
    #print(price_df)

    df = price_df




    while objStockChart.Continue:
        objStockChart.BlockRequest()
        len = objStockChart.GetHeaderValue(2)
        print("==============================================")
        value_list = []
        for i in range(len):
            day = objStockChart.GetDataValue(0, i)
            lock_info = objStockChart.GetDataValue(3, i)
            lock_before = objStockChart.GetDataValue(6, i)
            lock_after = objStockChart.GetDataValue(7, i)

            # 데이터 확인해보기
            #print(day, open, high, low, close, vol, vol_mon)
            value_list.append([day, lock_info, lock_before, lock_after ])



        price_df = pd.DataFrame(value_list)
        price_df.columns =["Date", "lock_info", "lock_before", "lock_after"]
        #print(price_df[['Date','num_listed', 'cap', 'ratio_foreigner', 'instit_net_buy','instit_net_budycum']])
        #print(price_df)
        df = pd.concat([df, price_df],axis=0, ignore_index=False)

    df['Date'] = pd.to_datetime(df['Date'].astype(str), format='%Y-%m-%d')
    df.set_index(['Date'], inplace=True)
    #print(f'-----------------total df : \n{df}')

    return df

#df = get_chart_daily(code=ticker, start_date="20000101")
#print(f'-----------------total df : \n \
#{df}')
# start_date="20000101" : yyyymmdd
#print(df.keys().tolist())
