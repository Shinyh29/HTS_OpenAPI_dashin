import win32com.client

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥

print("코스피 종목코드", len(codeList))


import pandas as pd
df_ks = pd.DataFrame()

codes = []
names = []

for i, code in enumerate(codeList):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)

    print(i, code, secondCode, stdPrice, name)
    codes.append(code)
    names.append(name)

df_ks['Ticker'] = codes
df_ks['Market'] = 'ks'
df_ks['Value'] = names





print("코스닥 종목코드", len(codeList))
df_kq = pd.DataFrame()
codes = []
names = []

for i, code in enumerate(codeList2):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)
    codes.append(code)
    names.append(name)

df_kq['Ticker'] = codes
df_kq['Market'] = 'kq'
df_kq['Value'] = names



print("거래소 + 코스닥 종목코드 ", len(codeList) + len(codeList2))

print(df_ks)
print(df_kq)

# ignore_index = True : reset index
dfs = pd.concat([df_ks, df_kq], axis = 0, ignore_index= True)
print(dfs)

### ------------df 2 aws-ec2-mysql
import pymysql
import time
from sqlalchemy import create_engine

item_tb = 'tickers'
pw ='0000'
ip_public = '3.35.27.15'
port = '3306'
db_name = 'ssiaat_shin'
engine = create_engine("mysql+pymysql://root:" + pw + f"@{ip_public}:{port}/{db_name}?charset=utf8",
                           encoding='utf-8')





try:
    dfs.to_sql(name='tickers', con=engine, if_exists='append', index=False)
except Exception as e:
    print(f'{e}_____Failed to bulkdf 2 EC2 insert')

