### 이것도  blockrequest 수로  요청해야함  일자 1000일 ~ 2000 (일간 넘어감 ) 가능,.
# http://cybosplus.github.io/cpsysdib_rtf_1_/cpsvr7254.htm
#  순매매 수량  ( 순매수금액  은 잘안받아지고 있음  )


import win32com.client
import pandas
import numpy


# 객체 생성
inCpSvr7254 = win32com.client.Dispatch("CpSysDib.CpSvr7254")

inCpSvr7254.SetInputValue(0, "A005930")
inCpSvr7254.SetInputValue(1, 6)
inCpSvr7254.SetInputValue(2, 20210101)
inCpSvr7254.SetInputValue(3, 20210304)
inCpSvr7254.SetInputValue(4, '0')
inCpSvr7254.SetInputValue(5, 0)   # 5 - (short)  투자자

inCpSvr7254.BlockRequest()

count = inCpSvr7254.GetHeaderValue(1)
print(f'count : {count}')


date_list = []
data_list1 = []
data_list2 = []
data_list3 = []


## 1 회
import pandas as pd
df = pd.DataFrame()
for i in range(count):
    # print("-----------------------------")
    date_list.append(inCpSvr7254.GetDataValue(0, i))
    data_list1.append(inCpSvr7254.GetDataValue(1, i))
    data_list2.append(inCpSvr7254.GetDataValue(2, i))
    data_list3.append(inCpSvr7254.GetDataValue(3, i))

df['Date'] = date_list
df['indiv'] = data_list1
df['foreign'] = data_list2
df['instit'] = data_list3



sum_count = 0




while inCpSvr7254.Continue:
    inCpSvr7254.BlockRequest()
    count = inCpSvr7254.GetHeaderValue(1)
    sum_count += count
    print(f'count : {sum_count}')
    if sum_count > 300:
        break;

    df = pd.DataFrame()
    for i in range(count):
        #print("-----------------------------")
        date_list.append( inCpSvr7254.GetDataValue(0, i) )
        data_list1.append( inCpSvr7254.GetDataValue(1, i) )
        data_list2.append( inCpSvr7254.GetDataValue(2, i) )
        data_list3.append( inCpSvr7254.GetDataValue(3, i) )

    df['Date'] = date_list
    df['indiv'] = data_list1
    df['foreign'] = data_list2
    df['instit'] = data_list3

print(df)