import os
import win32com.client as win32
import win32gui


##########################################################

# # 원격 프로시저 호출 (RPC) 문제/ 특히 반복적인 파일 열기/저장 과정에서 한글이 불안정

# # 작업 디렉토리 설정
# os.chdir("C:/Users/user/Desktop/phthon/2_hg/sub")

# # 한글 객체 생성
# hwp = win32.Dispatch("HwpFrame.HwpObject")

# # 한글 창 핸들 찾기
# hwnd = win32gui.FindWindow(None, "빈 문서 1 - 한글")
# print(hwnd)

# # 한글 창 숨기기
# win32gui.ShowWindow(hwnd, 0)

# # 파일 경로 검사 모듈 등록
# hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# BASE_DIR = "C:/Users/user/Desktop/phthon/2_hg/sub"

# for filename in os.listdir(BASE_DIR):
#     if filename.endswith(".hwp"):
#         hwp.Open(os.path.join(BASE_DIR, filename))
        
#         # PDF로 저장
#         hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
#         hwp.HParameterSet.HFileOpenSave.filename = os.path.join(BASE_DIR, filename.replace(".hwp", ".pdf"))
#         hwp.HParameterSet.HFileOpenSave.Format = "PDF"
#         hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)

###########################################################

# 한글을 껐다가 다시 켜서 작업하는것
import os
import win32com.client as win32

BASE_DIR = "C:/Users/user/Desktop/phthon/2_hg/sub"

for filename in os.listdir(BASE_DIR):
    if filename.endswith(".hwp"):
        hwp = win32.Dispatch("HwpFrame.HwpObject")
        
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(os.path.join(BASE_DIR, filename))
        
        # PDF로 저장
        hwp.HAction.GetDefault("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = os.path.join(BASE_DIR, filename.replace(".hwp", ".pdf"))
        hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        hwp.HAction.Execute("FileSaveAsPdf", hwp.HParameterSet.HFileOpenSave.HSet)
        
        hwp.Quit()