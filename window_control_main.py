import os
from tqdm import tqdm
import time

os.system('C:\\Anaconda3\\envs\\py37_32bit\\python.exe C:\\PycharmProjects\\dashin_api\\window_control.py')
time.sleep(1)
os.system('C:\\Anaconda3\\envs\\py37_32bit\\python.exe C:\\PycharmProjects\\dashin_api\\window_control.py')

for i in tqdm(range(0,100)):
    time.sleep(0.1)

os.system('C:\\Anaconda3\\envs\\py37_32bit\\python.exe C:\\PycharmProjects\\dashin_api\\window_control_main.py')