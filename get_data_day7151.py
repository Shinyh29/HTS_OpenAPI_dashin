### 이것도  blockrequest 수로  요청해야함  일자 1000일 ~ 2000 (일간 넘어감 ) 가능,.
# http://cybosplus.github.io/cpsysdib_rtf_1_/cpsvr7254.htm
#  순매수금액  도전
# http://money2.daishin.com/e5/mboard/ptype_basic/HTS_Plus_Helper/DW_Basic_Read_Page.aspx?boardseq=288&seq=230&page=1&searchString=&p=&v=&m=
# [7240 일자별 대차거래]
# -----------------------------150 개  이상으로는   안보이나봄

from datetime import datetime, timedelta
import win32com.client
import pandas
import numpy
import pandas as pd
pd.set_option('display.max_columns', 100)
import window_control
import os
#code = '005930'

item_tb = 'gongmado'
#table_nm = item_tb
# 순매수대금_외국인 : netbuy_foreign  who = 2, param = 7
# 순매수대금_기관계 : netbuy_instit   who = 3, param = 7
# 외국인비율 : rate_foreign who = 2, param = 8 ?



inCpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7151")

import date_set


def get_credit7151(code="A005930"):
    inCpSvr7254.SetInputValue(0, code)
    inCpSvr7254.SetInputValue(1, ord('y'))  # 1 - (char)구분 , { 'y':융자, 'd': 대주 }
    inCpSvr7254.SetInputValue(2, ord('1'))  #  2 - (char)구분 , { '1':결제일, '2': 매매일 }
    #count = inCpSvr7254.GetHeaderValue(0)

    dfs = pd.DataFrame()
    global out_max
    out_max = date_set.today(-1)
    out_trigger = 0.0

    inCpSvr7254.BlockRequest()
    # print( inCpSvr7254.GetDataValue(1, 10) )
    # print(f'count : {count}')

    # set First --------------------------Date pad
    df = pd.DataFrame()

    data_list = []
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
    data_list11 = []

    count = inCpSvr7254.GetHeaderValue(0)
    print(f'----------------GetHeaderValue , Count : {count}')
    for i in range(0, count):
        # print("-----------------------------")
        data_list.append(inCpSvr7254.GetDataValue(0, i)) ## 일자
        data_list1.append(inCpSvr7254.GetDataValue(1, i))  ## 종가
        data_list2.append(inCpSvr7254.GetDataValue(2, i))  ## 대비
        data_list3.append(inCpSvr7254.GetDataValue(3, i))  ## 등락율
        data_list4.append(inCpSvr7254.GetDataValue(4, i))  ## 거래량
        data_list5.append(inCpSvr7254.GetDataValue(5, i))  ## 신규
        data_list6.append(inCpSvr7254.GetDataValue(6, i))  ## 상환
        data_list7.append(inCpSvr7254.GetDataValue(7, i))  ## 잔고
        data_list8.append(inCpSvr7254.GetDataValue(8, i))  ## 금액
        data_list9.append(inCpSvr7254.GetDataValue(9, i))  ## 잔고대비
        data_list10.append(inCpSvr7254.GetDataValue(10, i))  ## 공여율
        data_list11.append(inCpSvr7254.GetDataValue(11, i))  ## 잔고율

    # -------------------------------------------INSERT

    cols = ['date', '종가', '대비', '등락율', '거래량', '신규', '상환',
            '잔고', '금액', '잔고대비','공여율','잔고율' ]



    df[f'{cols[0]}'] = data_list
    df[f'{cols[1]}'] = data_list1
    df[f'{cols[2]}'] = data_list2
    df[f'{cols[3]}'] = data_list3
    df[f'{cols[4]}'] = data_list4
    df[f'{cols[5]}'] = data_list5
    df[f'{cols[6]}'] = data_list6
    df[f'{cols[7]}'] = data_list7
    df[f'{cols[8]}'] = data_list8
    df[f'{cols[9]}'] = data_list9
    df[f'{cols[10]}'] = data_list10
    df[f'{cols[11]}'] = data_list11
    print(df)
    dfs = pd.concat([dfs, df], axis=0, ignore_index=False)

    print(f'========== df.iloc[0] date : {df.iloc[0].date}')
    out_trigger = df.iloc[0].date
    out_max = out_trigger
    out_trigger = 0
    print(f'out_trigger : {out_trigger}')
    print(f'out_max : {out_max}')


    while out_max != out_trigger:
        #
        # for idx, k in enumerate( range(0,count) ):
        #print(f'--------------------------idx / count : {idx}/ {count}')
        inCpSvr7254.BlockRequest()
        # print( inCpSvr7254.GetDataValue(1, 10) )
        # print(f'count : {count}')


        df = pd.DataFrame()

        data_list = []
        data_list1= []
        data_list2 = []
        data_list3=[]
        data_list4=[]
        data_list5=[]
        data_list6=[]
        data_list7=[]
        data_list8=[]
        data_list9=[]
        data_list10 = []
        data_list11 = []

        count = inCpSvr7254.GetHeaderValue(0)
        print(f'----------------GetHeaderValue , Count : {count}')
        for i in range(0 ,count):
            # print("-----------------------------")
            data_list.append(inCpSvr7254.GetDataValue(0, i))
            data_list1.append(inCpSvr7254.GetDataValue(1, i))  ## 일별 순매수 금액
            data_list2.append(inCpSvr7254.GetDataValue(2, i))  ## 일별 순매수 금액
            data_list3.append(inCpSvr7254.GetDataValue(3, i))  ## 일별 순매수 금액
            data_list4.append(inCpSvr7254.GetDataValue(4, i))  ## 일별 순매수 금액
            data_list5.append(inCpSvr7254.GetDataValue(5, i))  ## 일별 순매수 금액
            data_list6.append(inCpSvr7254.GetDataValue(6, i))  ## 일별 순매수 금액
            data_list7.append(inCpSvr7254.GetDataValue(7, i))  ## 일별 순매수 금액
            data_list8.append(inCpSvr7254.GetDataValue(8, i))  ## 일별 순매수 금액
            data_list9.append(inCpSvr7254.GetDataValue(9, i))  ## 일별 순매수 금액
            data_list10.append(inCpSvr7254.GetDataValue(10, i))  ## 공여율
            data_list11.append(inCpSvr7254.GetDataValue(11, i))  ## 잔고율
            #data_list10.append(inCpSvr7254.GetDataValue(10, i))  ## 일별 순매수 금액


        # -------------------------------------------INSERT
        cols = ['date', '종가', '대비', '등락율', '거래량', '신규', '상환',
                '잔고', '금액', '잔고대비', '공여율', '잔고율']

        df[f'{cols[0]}']=  data_list
        df[f'{cols[1]}'] = data_list1
        df[f'{cols[2]}']= data_list2
        df[f'{cols[3]}']=data_list3
        df[f'{cols[4]}']=data_list4
        df[f'{cols[5]}']=data_list5
        df[f'{cols[6]}']=data_list6
        df[f'{cols[7]}']=data_list7
        df[f'{cols[8]}']=data_list8
        df[f'{cols[9]}']=data_list9
        df[f'{cols[10]}'] = data_list10
        df[f'{cols[11]}'] = data_list11
        print(df)


        print(f'========== df.iloc[0] date : {df.iloc[0].date}')
        out_trigger = df.iloc[0].date
        #out_max = out_trigger
        print( f'out_trigger : {out_trigger}')
        if out_trigger > out_max:
            print('-----Loop OUT ! ')
            break
        dfs = pd.concat([dfs, df], axis=0, ignore_index=False)
        print( f'out_max : {out_max}')

    dfs= dfs.drop_duplicates(subset=['date'])
    dfs.reset_index(inplace=True, drop=True)
    return dfs

print(get_credit7151(code="A005930") )
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



