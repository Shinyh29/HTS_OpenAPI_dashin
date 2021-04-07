# Chart obj  의 item_tb ( keys ) 2 db ec2
# 이거 하나면 됨
import pymysql
from sqlalchemy import create_engine
import pandas as pd

# conn = pymysql.connect(host='3.35.27.15', port=3306, user='root', password='0000', db='ssiaat_shin',
# charset='utf8')
engine = create_engine("mysql+pymysql://root:" + "0000" + "@13.209.4.191:3306/ssiaat_shin?charset=utf8",
                           encoding='utf-8')
sql = f"""
SELECT * FROM ssiaat_shin.listed_cap;
"""
sql
#engine
conn = engine
conn
df_tickers = pd.read_sql(sql, con=conn)

## 시총 5000 억원 이상 -> 절반
df_cap5000 = df_tickers[df_tickers['cap'] > 5.00 * 10**11]  ### 상위 5천 억 이상 (1000종목)
import numpy as np
names = df_cap5000.name.values.tolist()
tickers = df_cap5000.Ticker.values.tolist()
print(len(tickers))
tickers

for idx in range(0, len(df_cap5000)):
    print(f'{idx}, {df_cap5000.name.iloc[idx]} , {df_cap5000.Ticker.iloc[idx]}'    )

def get_ticker2nm(ticker):
    get_name = df_tickers[df_tickers.Ticker == ticker ].name.values[0]
    return get_name

def get_nm2ticker(nm):
    get_name = df_tickers[df_tickers.name == nm ].Ticker.values[0]
    return get_name


def get_notin_dates(df, df2):
    isin_dates = []
    try:
        df = df.reset_index()
    except:
        None
    for unit_date in df.Date.tolist():
        if unit_date not in df2.Date.tolist():
            isin_dates.append(unit_date)

    temp_df = df
    temp_df.set_index(['Date'], inplace=True)
    temp_df = temp_df.loc[isin_dates].reset_index()
    return temp_df




import FinanceDataReader as fdr
df_etf_kr = fdr.EtfListing('KR')


df_etf_kr.Symbol = "A" + df_etf_kr.Symbol
etfs = df_etf_kr.Symbol.tolist()
stocks = []
idx = 0
for u in tickers:
    if u in etfs:
        idx += 1
        print(idx, u, get_ticker2nm(u))

    elif "Q" in u:
        idx += 1
        print(idx, u, get_ticker2nm(u))
    else:
        stocks.append(u)

print(f'len[stocks]  : {len(stocks)}')
print(stocks)

import numpy as np
import time
import StockChart_get_day


from tqdm import tqdm
import date_set

for idx, ticker in enumerate(stocks):
    unit_df = pd.DataFrame()
    df = StockChart_get_day.get_chart_daily(code=ticker, start_date= date_set.today(-5000) ) #'20000101')
    df = df.reset_index()
    df['Ticker'] = ticker

    table_nms = df.keys().tolist()
    for table_nm in table_nms:
        if table_nm in ['Ticker','Date','Open', 'High', 'Low',  'Volume']: #, 'netbuy_instit']:
            # 첫번째 호출  " 이유를 알수없는 Ticker NaN
            None
        else:
            #  가져오지 않고 있는것  얻기
            #  첫번째 Close  ->  NaN 으로 처리
            print(f' -----------------------table_nm : {table_nm}')
            print(f' -----------------------ticker : {ticker}')
            #time.sleep(3)


            unit_df = df[['Ticker','Date',table_nm]]
            print(unit_df)
            #unit_df = unit_df[['Ticker','Date',table_nm]]
            unit_df.rename(columns={table_nm: "Value"},inplace = True)
            unit_df.columns = unit_df.columns.to_series().apply(lambda x: x.strip() )
            #unit_df['Ticker'] = ticker
            #unit_df.rename(columns={table_nm : "Value"}, inplace=True)
            #unit_df['Value'] = df[f'{table_nm}']


            # ====================================================== #
            ### unit_df 2 EC2 db
            ### 근데  읽은다음에 없는 날짜만

            case_if = 'Close'
            if table_nm == case_if:
                # 첫번째 조회 테이블을 패스
                None
            else:
                try:
                    unit_df = unit_df.reset_index(drop=False)
                except:
                    None

                unit_df.columns = unit_df.columns.to_series().apply(lambda x: x.strip())
                #print(unit_df.tail())
                unit_df = unit_df[['Ticker', 'Date', 'Value']]
                #unit_df = unit_df.drop(['index'],axis=1)


                print('-----------------unit_df')
                print(unit_df.tail())
                # pk 따라서  Ticker, Date 겹치는것 있으면
                ## read_sql ascending True ->  과거값이 위로
                ## if

                sql = f"""
                SELECT * FROM ssiaat_shin.{table_nm} WHERE Ticker = '{ticker}' ORDER BY {table_nm}.Date DESC;
                """
                read_unit_df = pd.read_sql(sql, con=conn)

                print(f' -----------------------table_nm : {table_nm}')
                print('-----------------From EC2, read_unit_df')
                print(read_unit_df)
                # Date slice 해서 안넣고
                ## ----------unit_df.Date.tolist()
                ##  isin read_unit_df.Date.tolist()
                temp_df = get_notin_dates(unit_df, read_unit_df)
                temp_df = temp_df[['Ticker', 'Date', 'Value']]
                print(f' -----------------------table_nm : {table_nm}')
                print(f'--------------------will insert {ticker},  {get_ticker2nm(f"{ticker}")}')
                print(temp_df)

                # 겹치는것 없는 부분만 넣고
                # 시작이니까  다넣고

                # --------------------read,

                # --------------------insert : df 2 table
                try:
                    #None
                    temp_df.to_sql(name=f'{table_nm}', con=conn, if_exists='append', index=False)
                except Exception as e:
                    print(f'{e} ______ Failed to unit_df 2 EC2 insert')
                for i in tqdm(range(0, 1)):
                    time.sleep(0.1)

    for i in tqdm(range(0, 1)):  ## 100 에서   시간 너무오래걸려서  다시
        time.sleep(0.1)
