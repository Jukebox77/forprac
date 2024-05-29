import re
import win32com.client as win32


# 레지스트리 등록한 보안모듈
hwp = win32.gencache.EnsureDispatch("HwpFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
hwp.Open(r"C:/Users/user/Desktop/phthon/목차만들기.hwp")

title_list = []

hwp.InitScan()
while True:
    text = hwp.GetText()
    if text[0] == 0:
        hwp.ReleaseScan()
        break
    else:
        if re.match(r"[\d가-힣]+\.\s", text[1].strip()):
            hwp.MovePos(201, 0, 0)
            title_list.append(hwp.GetPos())
        else:
            pass

for titles in title_list:
    hwp.SetPos(*titles)
    hwp.Run("MarkTitle")

hwp.MovePos(2, 0, 0)
hwp.Run("BreakPage")
hwp.MovePos(2, 0, 0)