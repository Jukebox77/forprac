import win32com.client as win32
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# 파일 선택 창을 표시하기 위한 함수
def select_file(file_type):
    root = tk.Tk()
    root.withdraw()  # 루트 윈도우 숨기기
    if file_type == 'HWP':
        file_path = filedialog.askopenfilename(title="Select HWP file", filetypes=[("HWP files", "*.hwp")])
    elif file_type == 'Excel':
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return file_path

# HWP 파일 선택
hwp_file_path = select_file('HWP')

# 엑셀 파일 선택
excel_file_path = select_file('Excel')

# HWP 객체 생성
hwp = win32.Dispatch("HwpFrame.HwpObject")

# HWP 창 표시
hwp.XHwpWindows.Item(0).Visible = True

# 엑셀 파일 읽기 (역순으로 데이터프레임 뒤집기)
excel = pd.read_excel(excel_file_path)[::-1]
print(excel)

# 필요한 모듈 등록
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# HWP 파일 열기
hwp.Open(hwp_file_path, None, None)

# 필드 목록 가져오기
hwp1 = hwp.GetFieldList(1, 0)
cleaned_hwp1 = hwp1.replace("\x02", "")
print("Cleaned String:", cleaned_hwp1)

# 구분자 '{{0}}'를 기준으로 문자열을 나누어 리스트로 변환
field_list = cleaned_hwp1.split("{{0}}")
if field_list[-1] == "":
    field_list.pop()
print("Field List:", field_list)

# 현재 페이지 전체 선택 및 복사
hwp.Run("SelectAll")
hwp.Run("Copy")

# 엑셀 데이터프레임의 각 행을 처리
for index, row in excel.iterrows():
    # 필드에 값 입력
    for field in field_list:
        if field in row:
            hwp.PutFieldText(f"{field}{{{{}}}}", str(row[field]))
            # time.sleep(0.1)  # 필드 입력 후 지연 시간 추가

    # 문서 끝으로 이동하고 페이지 나누기 (Ctrl + Enter)
    hwp.Run("MoveDocEnd")
    hwp.Run("InsertPageBreak")

    # 첫 번째 페이지 붙여넣기
    hwp.Run("MoveDocBegin")
    hwp.Run("Paste")

    # 0.05초 대기
    # time.sleep(0.3)
