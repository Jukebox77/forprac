import os  # 파일 경로를 다루기 위한 모듈
from tkinter.filedialog import askopenfilenames  # 파일 선택창을 띄우기 위한 모듈

import win32com.client as win32  # 아래아한글을 열기 위한 모듈


def 한글_시작():
    """
    아래아한글을 시작하는 함수
    """
    hwp = win32.Dispatch("HwpFrame.HwpObject")  # 한/글 실행
    # hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행
    hwp.XHwpWindows.Item(0).Visible = True  # 한/글 프로그램 백그라운드 해제
    hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")  # 보안모듈 등록
    return hwp


def 파일선택():
    """
    파일선택 함수
    """
    filelist = askopenfilenames(title="자간을 조정할 한/글문서를 모두 선택해주세요.",
                                initialdir=os.getcwd(),
                                filetypes=[("한/글 파일", "*.hwp *.hwpx")])
    return filelist


def 현재선택영역_글자수():
    """
    자간자동조정 함수에서
    라인 끝에 걸쳐진 단어의
    앞뒤길이를 각각 계산하기 위함.
    """
    hwp.InitScan(option=None, Range=0xff, spara=None, spos=None, epara=None, epos=None)  # 선택한 범위 탐색시작
    _, text = hwp.GetText()  # 텍스트 추출
    hwp.ReleaseScan()  # 탐색종료
    return len(text)  # 추출한 텍스트의 글자수 리턴


def 자간자동조정():
    """
    모든 라인을 순회하면서
    끝에 걸쳐친 단어를 탐색함.

    잘린 단어의 앞이 길면
    라인 전체의 자간을 줄이고,

    잘린 단어의 뒤가 길면
    라인 전체의 자간을 늘임.

    한 줄 문단이 되거나
    걸쳐진 단어가 없으면 종료.
    """
    count = 0
    while True:
        hwp.Run("MoveLineEnd")  # 라인의 끝으로 이동해서
        hwp.Run("MoveSelWordBegin")  # 끝에 걸쳐진 단어의 앞부분만 선택
        if count >= 10:
            print("10% 이상 자간조정으로, 원상복구함")
            for _ in range(count): hwp.Run("Undo")
        앞부분길이 = 현재선택영역_글자수()  # 잘린 단어 앞부분 글자수 확인
        if 앞부분길이 == 0:  # 단어가 잘려있지 않으면 다음 라인으로 넘어감
            break
        hwp.Run("MoveSelWordEnd")  # 다음 라인으로 넘어간 부분 선택
        뒷부분길이 = 현재선택영역_글자수()  # 잘린 단어 뒷부분 글자수 확인
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


def 컨트롤_내부_자간조정():
    """
    표나 글상자 등 텍스트가 들어가는
    모든 영역의 자간을 조정하기 위함
    """
    area = 1  # 본문 외 영역(표, 각주미주, 글상자, 도형 등)
    while True:
        area += 1
        hwp.SetPos(area, 0, 0)  # 해당 영역으로 이동해서
        if hwp.GetPos()[0] == 0:  # 영역이동 중 본문으로 돌아오면
            break  # 작업끝.
        while True:
            시작위치 = hwp.GetPos()
            자간자동조정()  # 영역 첫 번째 라인 자간조정 하고,
            hwp.Run("MoveLineEnd")
            hwp.Run("MoveNextChar")  # 다음 라인으로 넘어감
            if hwp.GetPos()[0] != 0 and hwp.GetPos()[0] >= area:
                area = hwp.GetPos()[0]
            print(area)
            if hwp.GetPos() == 시작위치:
                break

        # area += 1  # 다음 영역으로 넘어감


def 끝위치추출():
    """
    본문 탐색 while문의 종료 조건으로
    "문서 끝에 도착하면 반복종료"를 구현하기 위해
    문서 끝 위치를 미리 추출해 둠
    """
    hwp.Run("MoveDocEnd")  # 문서 끝으로 이동한 후
    end_pos = hwp.GetPos()  # 문서 끝 위치(좌표) 저장
    hwp.Run("MoveDocBegin")  # 다시 문서 처음으로 이동
    return end_pos  # 저장한 좌표 리턴


if __name__ == '__main__':
    hwp = 한글_시작()  # 아래아한글 실행
    파일목록 = 파일선택()  # 자간 자동조정할 문서 전부 선택
    for 파일 in 파일목록:  # 문서 하나씩
        if 파일.endswith("x"):
            확장자 = "hwpx"
        else:
            확장자 = "hwp"
        hwp.Open(파일, Format=확장자.upper(), arg="")  # 한/글에서 열어서
        끝위치 = 끝위치추출()

        # 본문 자간조정
        while hwp.GetPos() != 끝위치:
            자간자동조정()
            hwp.Run("MoveLineEnd")
            hwp.Run("MoveNextChar")
        # 표 및 글상자 자간조정
        컨트롤_내부_자간조정()
        print("자간조정 작업 끝!")
        hwp.SaveAs(Path=hwp.Path.replace(f".{확장자}", f"(자간조정).{확장자}"), Format=hwp.XHwpDocuments.Item(0).Format, arg="")