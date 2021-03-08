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

    else:
        stocks.append(u)

print(f'len[stocks]  : {len(stocks)}')
print(stocks)


item_tb = 'netbuy_foreign'
import get_data_day7254_4

for idx, code in enumerate(stocks[4::]):
    print(f'-------------{idx}/{len(stocks)}: {code},  {get_ticker2nm(code)}')
    code = code.replace("A","")
    unit_df = get_data_day7254_4.get_netbuy( code = code, who=2 ,num =4300 )
    # let 4000일  == 20년치
    try:
        unit_df = unit_df.reset_index(drop=False)
    except:
        None
    try:
        unit_df = unit_df[['Ticker','Date','Value']]
    except:
        None
    print('-----------------unit_df')
    print(unit_df)
    # pk 따라서  Ticker, Date 겹치는것 있으면
    ## read_sql ascending True ->  과거값이 위로
    ## if

    sql = f"""
    SELECT * FROM ssiaat_shin.netbuy_foreign WHERE Ticker = 'A{code}' ORDER BY netbuy_foreign.Date DESC;
    """
    read_unit_df = pd.read_sql(sql, con=conn)

    print('-----------------From EC2, read_unit_df')
    print(read_unit_df)
    # Date slice 해서 안넣고
    ## ----------unit_df.Date.tolist()
    ##  isin read_unit_df.Date.tolist()
    temp_df = get_notin_dates(unit_df, read_unit_df)
    temp_df = temp_df[['Ticker','Date','Value']]
    print(f'--------------------will insert  A{code},  {get_ticker2nm(f"A{code}")}')
    print(temp_df)

    # 겹치는것 없는 부분만 넣고
    # 시작이니까  다넣고

    # --------------------read,



    # --------------------insert : df 2 table
    try:
        temp_df.to_sql(name = f'{item_tb}', con= conn, if_exists='append', index= False)
    except Exception as e:
        print(f'{e} ______ Failed to unit_df 2 EC2 insert')
