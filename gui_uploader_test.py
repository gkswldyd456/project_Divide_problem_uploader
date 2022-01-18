import os
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
from PIL import Image


root = Tk()
root.title("")


def browse_dest_path():
    folder_selected = filedialog.askdirectory()
    if folder_selected == "": # 사용자가 취소를 누를 때
        print("폴더 선택 취소")
        return
    txt_dest_path.delete(0, END)
    txt_dest_path.insert(0, folder_selected.replace("/", "\\"))
    global dir
    dir = folder_selected.replace("/", "\\")


def start():
    pass

# 저장 경로 프레임
path_frame = LabelFrame(root, text="작업폴더")
path_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_dest_path = Entry(path_frame, width=50)
txt_dest_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4) # 높이 변경

btn_dest_path = Button(path_frame, text="찾아보기", width=10, command=browse_dest_path)
btn_dest_path.pack(side="right", padx=5, pady=5)



# 문제집 번호
pronum_frame = LabelFrame(root, width=50, text="문제집 번호")
pronum_frame.pack(fill="both", padx=5, pady=5)

pronum_head = Text(pronum_frame, height=1)
pronum_head.pack(side="left", padx=5, pady=5, ipady=4)



# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start1 = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start1.pack(side="left", padx=5, pady=5)

# btn_start2 = Button(frame_run, padx=5, pady=5, text="수정해설저장", width=12, command=start_one_sol)
# btn_start2.pack(side="left", padx=5, pady=5)

# btn_start3 = Button(frame_run, padx=5, pady=5, text="수정문제저장", width=12, command=start_one_pro)
# btn_start3.pack(side="left", padx=5, pady=5)





root.resizable(False, False)
root.mainloop()

print(dir)
