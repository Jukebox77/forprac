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
    hwp.Open(os.path.join(os.getcwd(), i))
    sleep(0.5)
    
    hwp.PutFieldText("월당", 랜덤(1, 10))
    월당 = hwp.GetFieldText("월당")
    누계 += int(월당) if 월당.isdigit() else 0
    hwp.PutFieldText("월누", 누계)
    
    hwp.PutFieldText("화당", 랜덤(1, 10))
    화당 = hwp.GetFieldText("화당")
    누계 += int(화당) if 화당.isdigit() else 0
    hwp.PutFieldText("화누", 누계)
    
    hwp.PutFieldText("수당", 랜덤(1, 10))
    수당 = hwp.GetFieldText("수당")
    누계 += int(수당) if 수당.isdigit() else 0
    hwp.PutFieldText("수누", 누계)
    
    hwp.PutFieldText("목당", 랜덤(1, 10))
    목당 = hwp.GetFieldText("목당")
    누계 += int(목당) if 목당.isdigit() else 0
    hwp.PutFieldText("목누", 누계)
    
    hwp.PutFieldText("금당", 랜덤(1, 10))
    금당 = hwp.GetFieldText("금당")
    누계 += int(금당) if 금당.isdigit() else 0
    hwp.PutFieldText("금누", 누계)
    
    hwp.Save()
    sleep(0.5)
    hwp.Run("FieldClose")


startdate = dt.datetime(2024, 1, 1)
delta = 0
# 디렉토리의 파일들을 순회
for file in os.listdir():
    file_path = os.path.join(os.getcwd(), file)
    hwp.Open(file_path)
    for 요일 in "월화수목금토일":
        nextdate = startdate + dt.timedelta(days=delta)

        # 필드에 날짜 입력
        hwp.PutFieldText(요일, nextdate.strftime("%y. %m. %d.") + f"({'월화수목금토일'[nextdate.weekday()]})")
        delta += 1

        # 파일 저장
        hwp.Save()
        sleep(0.05)
    
    # 파일 닫기
    hwp.Run("FileClose")
    # 충분한 시간 대기
    # sleep(0.05)

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