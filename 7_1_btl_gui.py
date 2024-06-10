import tkinter.ttk as ttk
from tkinter import *

root=Tk()
root.title("Betweenletters GUI")

# 파일 프레임 (파일 추가, 선택 삭제)
file_frame = Frame(root)
file_frame.pack(fill="x", padx = 5, pady = 5)

btn_add_file = Button(file_frame, padx = 5, pady = 5, width = 12,  text="자간조정")
btn_add_file.pack(side="left")


root.resizable(False, False)
root.mainloop()