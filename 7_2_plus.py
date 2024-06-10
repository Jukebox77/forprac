import os  # 파일 경로를 다루기 위한 모듈
import tkinter.ttk as ttk
from tkinter import *
from tkinter.filedialog import askopenfilenames  # 파일 선택창을 띄우기 위한 모듈
import win32com.client as win32  # 아래아한글을 열기 위한 모듈

# 아래아한글 관련 함수들 (기존 코드 그대로)

def 한글_시작():
    hwp = win32.Dispatch("hwpframe.hwpobject")  # 한/글 실행
    hwp.XHwpWindows.Item(0).Visible = True  # 한/글 프로그램 백그라운드 해제
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안모듈 등록
    return hwp

def 파일선택():
    filelist = askopenfilenames(title="자간을 조정할 한/글문서를 모두 선택해주세요.",
                                initialdir=os.getcwd(),
                                filetypes=[("한/글 파일", "*.hwp *.hwpx")])
    return filelist

def 현재선택영역_글자수(hwp):
    hwp.InitScan(option=None, Range=0xff, spara=None, spos=None, epara=None, epos=None)  # 선택한 범위 탐색시작
    _, text = hwp.GetText()  # 텍스트 추출
    hwp.ReleaseScan()  # 탐색종료
    return len(text)  # 추출한 텍스트의 글자수 리턴

def 자간자동조정(hwp):
    count = 0
    while True:
        hwp.Run("MoveLineEnd")  # 라인의 끝으로 이동해서
        hwp.Run("MoveSelWordBegin")  # 끝에 걸쳐진 단어의 앞부분만 선택
        if count >= 15:
            print("15% 이상 자간조정으로, 원상복구함")
            for _ in range(count): hwp.Run("Undo")
        앞부분길이 = 현재선택영역_글자수(hwp)  # 잘린 단어 앞부분 글자수 확인
        if 앞부분길이 == 0:  # 단어가 잘려있지 않으면 다음 라인으로 넘어감
            break
        hwp.Run("MoveSelWordEnd")  # 다음 라인으로 넘어간 부분 선택
        뒷부분길이 = 현재선택영역_글자수(hwp)  # 잘린 단어 뒷부분 글자수 확인
        if not (앞부분길이 and 뒷부분길이):  # 한 줄 문단이면 넘어감
            hwp.Run("Cancel")  # 범위선택 해제
            hwp.Run("Cancel")  # 범위선택 해제
            break
        hwp.Run("MoveWordBegin")
        hwp.Run("MoveLineEnd")
        hwp.Run("MoveSelLineBegin")  # 라인 전체 선택해서
        if 앞부분길이 >= 뒷부분길이:  # 잘린 글자 앞부분이 길면?
            hwp.Run("CharShapeSpacingDecrease")  # 라인 자간 -1%
        else:  # 잘린 글자 뒷부분이 길면?
            hwp.Run("CharShapeSpacingIncrease")  # 라인 자간 +1%
        count += 1
        hwp.Run("Cancel")

def 컨트롤_내부_자간조정(hwp):
    area = 1  # 본문 외 영역(표, 각주미주, 글상자, 도형 등)
    while True:
        area += 1
        hwp.SetPos(area, 0, 0)  # 해당 영역으로 이동해서
        if hwp.GetPos()[0] == 0:  # 영역이동 중 본문으로 돌아오면
            break  # 작업끝.
        while True:
            시작위치 = hwp.GetPos()
            자간자동조정(hwp)  # 영역 첫 번째 라인 자간조정 하고,
            hwp.Run("MoveLineEnd")
            hwp.Run("MoveNextChar")  # 다음 라인으로 넘어감
            if hwp.GetPos()[0] != 0 and hwp.GetPos()[0] >= area:
                area = hwp.GetPos()[0]
            print(area)
            if hwp.GetPos() == 시작위치:
                break

def 끝위치추출(hwp):
    hwp.Run("MoveDocEnd")  # 문서 끝으로 이동한 후
    end_pos = hwp.GetPos()  # 문서 끝 위치(좌표) 저장
    hwp.Run("MoveDocBegin")  # 다시 문서 처음으로 이동
    return end_pos  # 저장한 좌표 리턴

def 자간조정():
    hwp = 한글_시작()  # 아래아한글 실행
    파일목록 = 파일선택()  # 자간 자동조정할 문서 전부 선택
    for 파일 in 파일목록:  # 문서 하나씩
        if 파일.endswith("x"):
            확장자 = "hwpx"
        else:
            확장자 = "hwp"
        hwp.Open(파일, Format=확장자.upper(), arg="")  # 한/글에서 열어서
        끝위치 = 끝위치추출(hwp)

        # 본문 자간조정
        while hwp.GetPos() != 끝위치:
            자간자동조정(hwp)
            hwp.Run("MoveLineEnd")
            hwp.Run("MoveNextChar")
        # 표 및 글상자 자간조정
        컨트롤_내부_자간조정(hwp)
        print("자간조정 작업 끝!")
        hwp.SaveAs(Path=hwp.Path.replace(f".{확장자}", f"(자간조정).{확장자}"), Format=hwp.XHwpDocuments.Item(0).Format, arg="")

# GUI 설정
def create_gui():
    root = Tk()
    root.title("Betweenletters GUI")
    root.geometry("500x300")
    root.configure(bg="#2c3e50")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", font=("Helvetica", 12), padding=10)
    style.configure("TLabel", font=("Helvetica", 14), padding=10, background="#2c3e50", foreground="#ecf0f1")

    # 프레임 설정
    frame = Frame(root, bg="#2c3e50")
    frame.pack(pady=20)

    # 제목 라벨
    lbl_title = Label(frame, text="Betweenletters GUI", bg="#2c3e50", fg="#ecf0f1", font=("Helvetica", 16, "bold"))
    lbl_title.pack(pady=10)

    # 설명 라벨
    lbl_desc = Label(frame, text="한글 문서의 자간을 자동으로 조정합니다.", bg="#2c3e50", fg="#bdc3c7", font=("Helvetica", 12))
    lbl_desc.pack(pady=5)

    # 버튼 추가
    btn_add_file = ttk.Button(frame, text="자간조정", command=자간조정)
    btn_add_file.pack(pady=20)

    root.resizable(False, False)
    root.mainloop()

# 실행
if __name__ == "__main__":
    create_gui()
