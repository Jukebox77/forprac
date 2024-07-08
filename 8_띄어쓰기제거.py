import os
import tkinter.ttk as ttk
from tkinter import *
from tkinter.filedialog import askopenfilename
import win32com.client as win32
import win32gui

def open_hwp():
    hwp = win32.Dispatch("HwpFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
    return hwp

def select_hwp_file():
    root = Tk()
    root.withdraw()  # 창을 숨기고
    root.attributes('-topmost', True)  # 창을 최상위로 유지

    file_path = askopenfilename(filetypes=[("HWP Files", "*.hwp")])

    root.destroy()  # 윈도우 창 파괴
    return file_path

def read_hwp_content(hwp, file_path):
    if file_path:
        hwp.Open(file_path)
        hwp.Run("SelectAll")  # 전체 선택
        hwp.Run("Copy")  # 복사 동작 실행
        content = hwp.GetTextFile("TEXT", "").strip()  # TEXT 형식의 전체 텍스트 가져오기
        
        # 객체에 None 할당하여 해제
        hwp = None
        
        return content
    else:
        return "No file selected"

def main():
    file_path = select_hwp_file()
    if not file_path:
        print("No file selected.")
        return

    hwp = open_hwp()
    content = read_hwp_content(hwp, file_path)
    
    # 각 줄의 맨 앞 불필요한 띄어쓰기 제거
    cleaned_content = "\n".join(line.lstrip() for line in content.splitlines())
    
    # 결과 출력
    print("Cleaned Content:")
    print(cleaned_content)

if __name__ == "__main__":
    main()
