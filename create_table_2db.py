import pymysql

import time
from sqlalchemy import create_engine

pw ='0000'
ip_public = '13.209.4.191'
port = '3306'
db_name = 'ssiaat_shin'

engine = create_engine("mysql+pymysql://root:" + pw + f"@{ip_public}:{port}/{db_name}?charset=utf8",
                           encoding='utf-8')

item_tb = 'lock_info'
# 기관 netbuy_instit
# 개인 netbuy_indiv
# 외국인 netbuy_foreign
# 외국인보유비율 rate_foreign

def TableCreater(item_tb):
    print(f''' make Table''')
    conn = pymysql.connect(host=ip_public, port=3306, user='root', password=pw, db=db_name,
                           charset='utf8')
    with conn.cursor() as curs:
        sql = f"""
        CREATE TABLE {item_tb} (
        Ticker VARCHAR(30) NOT NULL,
        Date Date NOT NULL,
        Value VARCHAR(128) NOT NULL,
        PRIMARY KEY(Ticker, Date)
        );
        """
        #
        # Date DATE NOT NULL,
        curs.execute(sql)
        print(f'{sql}')
        conn.commit()
        #codes = dict()
        conn.close()


#
TableCreater(item_tb)
