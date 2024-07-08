import win32com.client as win32
import pandas as pd

# HWP 객체 생성
hwp = win32.Dispatch("HwpFrame.HwpObject")

# HWP 창 표시
hwp.XHwpWindows.Item(0).Visible = True

# 엑셀 파일 읽기 (역순으로 데이터프레임 뒤집기)
excel = pd.read_excel("C:/Users/user/Desktop/phthon/2_hg/defect_inspection/result_3_2.xlsx")[::-1]
print(excel)

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

# 엑셀 데이터프레임의 각 행을 처리
for index, row in excel.iterrows():
    # 첫 번째 페이지 붙여넣기 (맨 처음은 복사된 페이지가 이미 있으므로 제외)
    if index > 0:
        hwp.Run("MoveDocEnd")
        hwp.Run("Paste")

    # 필드에 값 입력
    for field in field_list:
        if field in row:
            hwp.PutFieldText(f"{field}{{{{}}}}", str(row[field]))
    
    # 다음 페이지를 추가 (마지막 행 제외)
    if index < len(excel) - 1:
        hwp.Run("MoveDocEnd")
        hwp.Run("InsertPageBreak")
