import os
import shutil
import re
import win32com.client as win32
import win32gui
from time import sleep
from random import randint as 랜덤
import datetime as dt

# 레지스트리 등록한 보안모듈
hwp = win32.gencache.EnsureDispatch("HwpFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwnd = win32gui.FindWindow(None, "빈 문서 1 - 한글")
win32gui.ShowWindow(hwnd, 0)

os.chdir(r"C:/Users/user/Desktop/phthon/2_hg/sub")
file = os.listdir()[0]

# print(int(re.findall(r"\d+", file)[1]))

############################################################
# 한글파일 생성
# original = file.split("(")[0] + "(01주차).hwp"
# if file != original:
#     shutil.copy(file, original)

# for i in range(2,1+52):
#     shutil.copy(file, file.split("(")[0] + f"({i:02}주차).hwp")
############################################################

# 몇주차 입력
# for i in os.listdir():
#     주차 =  re.findall(r"\d+", i)[1]
#     hwp.Open(os.path.join(os.getcwd(), i))
#     sleep(0.05)
#     hwp.PutFieldText('몇주차', int(주차))
#     hwp.Save()
#     sleep(0.05)
#     print(hwp.GetFieldText("몇주차"))
#     hwp.Run("FileClose")

############################################################

누계 = 0
for i in os.listdir():
    hwp.Open(os.path.join(os.getcwd()), i)
    sleep(0.5)
    hwp.PutFieldText("월당", str(랜덤(1, 10)))
    누계 += int(hwp.GetFieldText("월당"))
    hwp.PutFieldText("월누", "누계")
    
    hwp.PutFieldText("화당", str(랜덤(1, 10)))
    누계 += int(hwp.GetFieldText("화당"))
    hwp.PutFieldText("화누", "누계")
    
    hwp.PutFieldText("수당", str(랜덤(1, 10)))
    누계 += int(hwp.GetFieldText("수당"))
    hwp.PutFieldText("수누", "누계")
    
    hwp.PutFieldText("목당", str(랜덤(1, 10)))
    누계 += int(hwp.GetieldText("목당"))
    hwp.PutFieldText("목누", "누계")
    
    hwp.PutFieldText("금당", str(랜덤(1, 10)))
    누계 += int(hwp.GetFieldText("금당"))
    hwp.PutFieldText("금누", "누계")
    hwp.Save()
    sleep(0.5)
    hwp.Run("FieldClose")

############################################################

# 날짜 설정

# startdate = dt.datetime(2024,1,1)
# delta = 0

# for file in os.listdir():
#     hwp.Open(os.path.join(os.getcwd(), file))
#     for 요일 in "월화수목금토일":
#         date = startdate + dt.timedelta(days=delta)
#         hwp.PutFieldText(요일, date.strftime("%y. %m. %d.") + f"{'월화수목금토일'[date.weekday()]}")
#         delta += 1
#         hwp.Save()
#         sleep(0.05)
#         print(hwp.GetFieldText("요일"))

#     hwp.Run("FileClose")