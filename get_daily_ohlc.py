from pandas import DataFrame
import win32com.client
import win32com

# ### stockChart ###
def subStockChart(m_Code, m_NumHist):                                   ### 1
# def subStockChart(m_Code, m_FromDate, m_ToDate):                           ### 2
    stockChart = win32com.client.Dispatch("CpSysDib.StockChart")
    stockChart.SetInputValue(0, m_Code)    # 종목 코드
    stockChart.SetInputValue(1, ord('2'))    # 요청 구분 (개수로 요청)  ### 1
    # stockChart.SetInputValue(1, ord('1'))    # 요청 구분 (날짜로 요청)      ### 2
    # stockChart.SetInputValue(2, m_ToDate)  # 시작일자                       ### 2
    # stockChart.SetInputValue(3, m_FromDate)    # 끝일자                     ### 2
    stockChart.SetInputValue(4, m_NumHist)     # 요청개수                ### 1
    stockChart.SetInputValue(5, [0,2,3,4,5]) # 날짜, 시가, 고가, 저가, 종가
    stockChart.SetInputValue(6, ord('D'))    # 차트 구분 (일)

    # ## 데이터 호출
    stockChart.BlockRequest()
    num = stockChart.GetHeaderValue(3)
    data=[]
    for i in range(num):
        tempData ={}
        tempData['Date']=(stockChart.GetDataValue(0,i))
        tempData['Open']=float(format(stockChart.GetDataValue(1,i), '.2f')) # 선물값이 오류수정
        tempData['High']=float(format(stockChart.GetDataValue(2,i), '.2f'))
        tempData['Low']=float(format(stockChart.GetDataValue(3,i), '.2f'))
        tempData['Close']=float(format(stockChart.GetDataValue(4,i), '.2f'))
        data.append(tempData)

    # ## dataframe으로 전환
    resultDf = DataFrame(data, columns=['Date', 'Open', 'High', 'Low', 'Close'])
    return resultDf

from pandas import DataFrame
#from subDS import subStockChart # 위에 작성한 subStockChart 불러오기

if __name__ == "__main__":
    ### 자료가져오기
    code ='A005930' # 삼성전자 종목코드
    numHist =500 # 과거 100일치 데이터
    df=subStockChart(code, numHist)
    print(df)