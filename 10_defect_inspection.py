import win32com.client as win32
import pandas as pd

# HWP 객체 생성
hwp = win32.Dispatch("HwpFrame.HwpObject")

# HWP 창 표시
hwp.XHwpWindows.Item(0).Visible = True

# 엑셀 파일 읽기
excel = pd.read_excel("C:/Users/user/Desktop/phthon/2_hg/defect_inspection/result_3_2.xlsx")[::-1]  # 역순으로 데이터프레임 뒤집기

# print(excel)

# 필요한 모듈 등록
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# HWP 파일 열기
hwp.Open(r"C:/Users/user/Desktop/phthon/2_hg/defect_inspection/result_3_1.hwp", None, None)

# 필드 목록 가져오기
hwp1 = hwp.GetFieldList(1, 0)
cleaned_hwp1 = hwp1.replace("\x02", "")
print("Cleaned String:", cleaned_hwp1)

# 구분자 '{{0}}'를 기준으로 문자열을 나누어 리스트로 변환
field_list = cleaned_hwp1.split("{{0}}")
if field_list[-1] == "":
    field_list.pop()
print("Field List:", field_list)

# 첫 번째 페이지 복사하기
hwp.Run("SelectAll")
hwp.Run("Copy")

# award.xlsx의 각 행을 처리하고 award.hwp에 페이지를 추가하여 값을 입력
for index, row in excel.iterrows():
    # award.xlsx의 현재 행 데이터를 award.hwp에 입력
    for field in field_list:
        hwp.PutFieldText(f"{field}{{{{}}}}", str(row[field]))
    
    # 다음 페이지를 추가
    hwp.Run("MoveDocEnd")
    hwp.Run("InsertPageBreak")

    # 첫 번째 페이지 붙여넣기
    hwp.Run("MoveDocBegin")
    hwp.Run("Paste")
'''
# 작업이 끝난 후 첫 번째 페이지만 선택
hwp.Run("MoveDocBegin")    # 첫 페이지로 이동
hwp.Run("Select")          # 첫 페이지 선택

# 선택된 페이지 삭제
hwp.Run("TableDelete")     # 선택된 페이지 삭제'''