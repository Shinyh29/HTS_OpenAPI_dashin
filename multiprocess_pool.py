# try multiprocess in python
#!/usr/bin/env python
import subprocess


## 공매도 거래대금  종목별  insert

# start all programs
a = f'python C:\\PycharmProjects\\dashin_api\\window_control_main.py'
b= f'python C:\\PycharmProjects\\dashin_api\\get_data_day7238Tomysql.py'
c= f'python C:\\PycharmProjects\\dashin_api\\listedcap_netbuy2mysql.py'
processes = [subprocess.Popen(program) for program in [a, b,c]]
# wait
for process in processes:
    process.wait()