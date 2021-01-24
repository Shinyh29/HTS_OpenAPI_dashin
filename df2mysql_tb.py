import pandas as pd
import pymysql

import time
from sqlalchemy import create_engine

pw ='0000'
ip_public = '3.35.27.15'
port = '3306'
db_name = 'ssiaat_shin'

engine = create_engine("mysql+pymysql://root:" + pw + f"@{ip_public}:{port}/{db_name}?charset=utf8",
                           encoding='utf-8')

item_tb = 'tickers'


def TableCreater(item_tb):
    print(f''' make Table''')
    conn = pymysql.connect(host=ip_public, port=3306, user='root', password=pw, db=db_name,
                           charset='utf8')
    with conn.cursor() as curs:
        sql = f"""
        CREATE TABLE {item_tb} (
        Ticker VARCHAR(30) NOT NULL,
        Market VARCHAR(3) NOT NULL
        Value VARCHAR(64) NOT NULL,
        );
        """
        #PRIMARY KEY(Ticker, Date)
        # Date DATE NOT NULL,
        curs.execute(sql)
        print(f'{sql}')
        conn.commit()
        codes = dict()
        conn.close()


# def UpdateTable(item_tb):
#     print(f'{will_insert_rows.Date.iloc[0]}  -----will INSERT')
#     print(f''' �ㅻ쪟�덈뒗 媛�_row�ㅻ쭔 ��  ��젣 �� : ��뼱�뚯슦湲�''')
#     conn = pymysql.connect(host='13.125.133.11', port=3306, user='root', password='expo0407', db='exp_db',
#                            charset='utf8')
#     with conn.cursor() as curs:
#         sql = f'''
#         DELETE FROM {item_tb}
#         WHERE {item_tb}.Date = '{str(will_insert_rows['Date'].iloc[0])}';
#         '''
#
#         curs.execute(sql)
#         print(f'{sql}')
#         conn.commit()
#         codes = dict()
#         conn.close()

    # try:
    #     will_insert_rows.to_sql(name=item_tb, con=engine, if_exists='append', index=False)
    # except Exception as e:
    #     print(f'{e}_____Failed to bulkdf 2 EC2 insert')
    #     None

try:
    TableCreater(item_tb)
except Exception as e:
    print(f'{e}')