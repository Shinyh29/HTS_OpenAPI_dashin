### 이것도  blockrequest 수로  요청해야함  일자 1000일 ~ 2000 (일간 넘어감 ) 가능,.
# http://cybosplus.github.io/cpsysdib_rtf_1_/cpsvr7254.htm
#  순매수금액  도전

from datetime import datetime, timedelta
import win32com.client
import pandas
import numpy
import pandas as pd
pd.set_option('display.max_columns', 100)


#code = '005930'

def get_netbuy(code, who ,num):
    # code : 티커
    # who : 투자자
    # num : 일자
    Ticker = 'A' + code
    # 객체 생성
    inCpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7254")
    inCpSvr7254.SetInputValue(0, Ticker)
    inCpSvr7254.SetInputValue(1, 6)
    inCpSvr7254.SetInputValue(2, 20210101)
    inCpSvr7254.SetInputValue(3, 20210304)
    inCpSvr7254.SetInputValue(4, '0')

    inCpSvr7254.SetInputValue(5, who)   # 5 - (short)  투자자

    inCpSvr7254.BlockRequest()

    count = inCpSvr7254.GetHeaderValue(1)
    #print(f'count : {count}')


    date_list = []
    data_list1 = []
    data_list2 = []
    data_list3 = []
    data_list4 = []
    data_list5 = []
    data_list6 = []
    data_list7 = []
    data_list8 = []
    data_list9 = []
    data_list10 = []

    ## 1 회


    df = pd.DataFrame()
    for i in range(count):
        # print("-----------------------------")
        date_list.append(inCpSvr7254.GetDataValue(0, i))
        data_list10.append(inCpSvr7254.GetDataValue(10, i))  ## 일별 순매수 금액

    sum_count = 0




    while inCpSvr7254.Continue:
        inCpSvr7254.BlockRequest()
        count = inCpSvr7254.GetHeaderValue(1)
        sum_count += count
        print(f'count : {sum_count}')
        if sum_count > num:
            break;

        df = pd.DataFrame()
        for i in range(count):
            #print("-----------------------------")
            date_list.append( inCpSvr7254.GetDataValue(0, i) )
            data_list10.append(inCpSvr7254.GetDataValue(10, i))

        df['Date'] = date_list
        df['netbuy_foreign'] = data_list10
    df['Ticker'] = Ticker
    df = df[['Ticker', 'Date', 'netbuy_foreign']]
    df['Date']= apply(lambda x: pd.to_datetime(str(x), format='%Y%m%d'))

    #print(pd.to_datetime(df['Date'], format= '%Y-%m-%d') )

    return df


# < 그 이외의 경우 >
# 0 - (long) 일자
# 1 - (long) 매도수량
# 2 - (double) 매도수량비중
# 3 - (long) 매도금액(백만)
# 4 - (double) 매도금액비중
# 5 - (long) 매수수량
# 6 - (double) 매수수량비중
# 7 - (long) 매수금액(백만)
# 8 - (double) 매수금액비중
# 9 - (long) 일별순매수수량
# 10 - (long) 일별순매수금액(백만)


### 이걸  이제  sql 으로 집어넣자  대상종목은  5천억 이상 500여종목
"""
who 
5 - (short)  투자자

코드

내용

0

전체

1

개인

2

외국인

3

기관계

4

금융투자

5

보험

6

투신

7

은행

8

기타금융

9

연기금

10

국가지자체

11

기타외인

12

사모펀드

13

기타법인



"""

print(get_netbuy( code = '005930', who=2 ,num = 19))