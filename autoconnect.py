# ch08 aouto Connect.py
from pywinauto import application
import os
import time

os.system('taskkill /IM ncStarter* /F /T')
os.system('taskkill /IM CpStart* /F /T')
os.system('taskkill /IM DibServer* /F /T')

# wmic : window 시스템 정보를 조회, 변경 강제종료 신호 받으면 확인창 띄워서 한번다 프로세스 종료함.
os.system('wmic process where "name like \'%ncStarter%'" call terminate")
os.system('wmic process where "name like \'%CpStart%'" call terminate")
os.system('wmic process where "name like \'%DibServer%'" call terminate")

time.sleep(5)
app = application.Application()

# creon 프로그램 coStarter.exe
# 크레온 플러스 모드 /prj:cp 로 자동시작.
# id, pwd, 공인인증서 암호  실행인수로 지정,  후  자동 입력
app.start('C:\\DAISHIN\\STARTER\\ncStarter.exe /prj:cp /id:아이디 /pwd:계좌패스워드 /pwdcert:공동인증서패스워드 /autostart')