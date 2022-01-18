import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
import os, shutil

import win32com.client as win
import time
from PIL import Image
import olefile
import re


dir = r"C:\Users\HanJiYong\Desktop\Testhaha\[1]하하하하"


list_pro_hml_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[1문제] [1hml]" in i] # dir 중 "[1문제] [1hml]" 있는 파일 제목들(no 경로)  
print(list_pro_hml_files)
print(len(list_pro_hml_files))


list_pro_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[1문제] [2png]" in i] # dir 중 "[1문제] [2png]" 있는 파일 제목들(no 경로) 
list_sol_hml_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[2해설] [1hml]" in i] # dir 중 "[1문제] [1hml]" 있는 파일 제목들(no 경로)  
list_sol_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[2해설] [2png]" in i] # dir 중 "[2해설] [2png]" 있는 파일 제목들(no 경로) 
list_presol_hwp_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[3정답] [1hwp]" in i] # dir 중 "[3정답] [1hwp]" 있는 파일 제목들(no 경로)
list_presol_types = [i.split('.')[-2] for i in list_presol_hwp_files] # list_presol_hwp_files 중 정답 타입 부분만 
list_presol_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[3정답] [2png]" in i] # dir 중 "[3정답] [2png]" 있는 파일 제목들(no 경로)
list_common_hml_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[0공통문제] [1hml]" in i] # dir 중 "[0공통문제] [1hml]" 있는 파일 제목들(no 경로)  
list_common_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[0공통문제] [2png]" in i] # dir 중 "[0공통문제] [2png]" 있는 파일 제목들(no 경로)  
list_common_first_last = [(i.split('번 [0공통문제]')[0][-7:].split('-')[0], i.split('번 [0공통문제]')[0][-7:].split('-')[1]) for i in list_common_hml_files] # [공통문제]중 (시작번호, 끝번호) 튜플형태로 추출


presol_text = [] # list_presol_hwp_files 정답 파일 안의 텍스트들만 리스트화
for list_presol_hwp_file in list_presol_hwp_files:
    f = olefile.OleFileIO(os.path.join(dir, list_presol_hwp_file)) # HWP 파일 열기
    encoded_text = f.openstream('PrvText').read() # PrvText 스트림의 내용 꺼내기
    decoded_text = encoded_text.decode('UTF-16') # 유니코드를 UTF-16으로 디코딩
    re_decoded_text = re.sub('\n|\r', "", decoded_text)
    print(re_decoded_text)
    presol_text.append(re_decoded_text)
    f.close()

print(presol_text)

if len(list_common_hml_files) != 0: # 공통부분 문제번호 리스트 된 것 reverse 
    reverse_list_common_first_last = list_common_first_last 
    reverse_list_common_first_last.reverse()
    reverse_list_common_hml_files = list_common_hml_files
    reverse_list_common_hml_files.reverse()
    reverse_list_common_png_files = list_common_png_files
    reverse_list_common_png_files.reverse()

if len(list_common_hml_files) != 0:
    for idx, (first, last) in enumerate(reverse_list_common_first_last) : # 역순으로 공통부분 첫 문제번호 위치에 빈 텍스트 각 리스트에 추가
        list_pro_hml_files.insert(int(first)-1, reverse_list_common_hml_files[idx])
        list_pro_png_files.insert(int(first)-1, reverse_list_common_png_files[idx])
        list_sol_hml_files.insert(int(first)-1, "")
        list_sol_png_files.insert(int(first)-1, "")
        list_presol_hwp_files.insert(int(first)-1, "")
        list_presol_types.insert(int(first)-1, "")
        list_presol_png_files.insert(int(first)-1, "")
        presol_text.insert(int(first)-1, "")
    
print(list_pro_hml_files)
print(list_pro_png_files)
# print(list_sol_hml_files)
# print(list_sol_png_files)
# print(list_presol_hwp_files)
# print(list_presol_types)
# print(list_presol_png_files)
# print(list_common_hml_files)
# print(list_common_png_files)
# print(presol_text)


num=1 
idxnum = num - 1 
print(type(num))
print(type(idxnum))