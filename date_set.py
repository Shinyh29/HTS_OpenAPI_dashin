import pandas as pd
from pandas.tseries.offsets import BDay, Day
backday = 0

def today(backday):
    today = pd.datetime.today()
    what_date1 = today + BDay(backday)
    what_date1 = format(what_date1,'%Y%m%d')
    today = what_date1
    return today

#print(today(-1))