### 이것도  blockrequest 수로  요청해야함  일자 1000일 ~ 2000 (일간 넘어감 ) 가능,.
# http://cybosplus.github.io/cpsysdib_rtf_1_/cpsvr7254.htm
#  순매수금액  도전

from datetime import datetime, timedelta
import win32com.client
import pandas
import numpy
import pandas as pd
pd.set_option('display.max_columns', 100)
import window_control
import os
#code = '005930'

item_tb = 'rate_foreign'
#table_nm = item_tb
# 순매수대금_외국인 : netbuy_foreign  who = 2, param = 7
# 순매수대금_기관계 : netbuy_instit   who = 3, param = 7
# 외국인비율 : rate_foreign who = 2, param = 8 ?


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
    inCpSvr7254.SetInputValue(3, 20210308)
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
        # data_list1.append(inCpSvr7254.GetDataValue(1, i))  ## 일별 순매수 금액
        # data_list2.append(inCpSvr7254.GetDataValue(2, i))  ## 일별 순매수 금액
        # data_list3.append(inCpSvr7254.GetDataValue(3, i))  ## 일별 순매수 금액
        # data_list4.append(inCpSvr7254.GetDataValue(4, i))  ## 일별 순매수 금액
        # data_list5.append(inCpSvr7254.GetDataValue(5, i))  ## 일별 순매수 금액
        # data_list6.append(inCpSvr7254.GetDataValue(6, i))  ## 일별 순매수 금액
        # data_list7.append(inCpSvr7254.GetDataValue(7, i))  ## 일별 순매수 금액
        # data_list8.append(inCpSvr7254.GetDataValue(8, i))  ## 일별 순매수 금액
        # data_list9.append(inCpSvr7254.GetDataValue(9, i))  ## 일별 순매수 금액
        #data_list10.append(inCpSvr7254.GetDataValue(10, i))  ## 일별 순매수 금액
        data_list10.append(inCpSvr7254.GetDataValue(10, i))  ## 일별 순매수 금액

    sum_count = 0




    while inCpSvr7254.Continue:
        inCpSvr7254.BlockRequest()
        count = inCpSvr7254.GetHeaderValue(1)
        sum_count += count
        print(f'count : {sum_count}')
        window_control.close_window_titlenm(titlenm='CPSYSDIB')
        if sum_count > num:
            break;

        df = pd.DataFrame()
        for i in range(count):
            #print("-----------------------------")
            date_list.append( inCpSvr7254.GetDataValue(0, i) )
            # data_list1.append(inCpSvr7254.GetDataValue(1, i))
            # data_list2.append(inCpSvr7254.GetDataValue(2, i))
            # data_list3.append(inCpSvr7254.GetDataValue(3, i))
            # data_list4.append(inCpSvr7254.GetDataValue(4, i))
            # data_list5.append(inCpSvr7254.GetDataValue(5, i))
            # data_list6.append(inCpSvr7254.GetDataValue(6, i))
            # data_list7.append(inCpSvr7254.GetDataValue(7, i))
            # data_list8.append(inCpSvr7254.GetDataValue(8, i))
            # data_list9.append(inCpSvr7254.GetDataValue(9, i))
            data_list10.append(inCpSvr7254.GetDataValue(10, i))

            #data_list10.append(inCpSvr7254.GetDataValue(10, i)) 일별순매수금액

        df['Date'] = date_list
        df[f'{item_tb}'] = data_list10



    df['Ticker'] = Ticker
    df = df[['Ticker', 'Date', f'{item_tb}']]

    df.rename(columns={f'{item_tb}': "Value"}, inplace=True)
    df = df.iloc[1::]  # 하루전날 부터 받음  : Date 제대로 나오지 않음
    try:
        df['Date']= pd.to_datetime(df['Date'].astype(str), format='%Y-%m-%d')
    except Exception as e:
        print(f'{e}-----------Date_convert Err')
    #print(f'--------------table_nm : {item_tb}')
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
0전체
1개인
2외국인
3기관계
4금융투자
5보험
6투신
7은행
8기타금융
9연기금
10국가지자체
11기타외인
12사모펀드
13기타법인



"""

print(get_netbuy( code = '005930', who=2 ,num = 30))