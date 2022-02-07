import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
import time
import os
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.alert import Alert
import olefile
import re





def typo_open():
    url = 'https://typo.postmath.co.kr/Web2/WorkbookList.aspx' # 타이포 서버
    driver.execute_script(f'window.open("{url}");')
    # driver.get(url) # 페이지 열어라 
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(0.1)
    try:
        driver.find_element_by_css_selector('#mb_id').clear()
        driver.find_element_by_css_selector('#mb_id').send_keys('postmath21') # 포메 아이디 입력
        driver.find_element_by_css_selector('#mb_password').send_keys('wD5V8Wo9@efHe$jW') # 포메 비번 입력
        if driver.find_element_by_xpath('//*[@id="id_save"]').get_attribute('checked') == None : # 아이디 저장 체크 되어있는지 판단 (안눌려있으면 None, 눌려있으면 True) 
            driver.find_element_by_xpath('//*[@id="id_save"]').click() # -> 안눌려 있다면 눌러라
        driver.find_element_by_css_selector('#btnLogin').click() # 로그인버튼 눌러
    except NoSuchElementException:
        pass
    time.sleep(0.2)
    # typonum = int(pyautogui.prompt(text='문제집 번호(파일 번호)를 쓰시오.')) # 문제집번호
    re_url = f'https://typo.postmath.co.kr/Web2/WorkbookDetailList.aspx?seq={typonum}'
    driver.get(re_url) # 페이지 열어라 
    driver.find_element_by_css_selector('#ddlCount > option:nth-child(4)').click() # 500개 조회 선택


def check_exists_by_css_selector(css_selector): # css_selector로 요소 존재 여부 판단하기 
    try:
        driver.find_element_by_css_selector(css_selector)
    except NoSuchElementException:
        return False
    return True


def pro_upload(num): # 문제 업로드 
    idxnum = num -1 
    driver.implicitly_wait(0.5)
    while check_exists_by_css_selector('#cboxLoadedContent > iframe') ==False: # iframe 창이 아직 없으면 계속 눌러라
        try:
            driver.find_element_by_css_selector(f'#container > div > table.tb_base > tbody > tr:nth-child({num}) > td:nth-child(17) > a').click() # 1번 문제 업로드
            if check_exists_by_css_selector('#cboxLoadedContent > iframe') ==True:
                break
        except NoSuchElementException:
            pass
    driver.implicitly_wait(3)
    
    time.sleep(0.3)

    el = driver.find_element_by_class_name('cboxIframe') # iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
    driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨
    
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys('{0}'.format(list_pro_hml_files[idxnum])) # 문제 hml업로드
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload2"]').send_keys('{0}'.format(list_pro_png_files[idxnum])) # 문제 png업로드
    time.sleep(0.1)
    driver.find_element_by_css_selector('#btnSave').click()
    time.sleep(0.1)

    driver.switch_to.alert
    Alert(driver).accept()
    time.sleep(0.1)
    Alert(driver).accept()

    driver.switch_to.default_content() # 처음 frame으로 돌아가기
    time.sleep(0.1)




def sol_upload(num): # 해설 업로드 
    idxnum = num -1 
    driver.implicitly_wait(0.5)
    while check_exists_by_css_selector('#cboxLoadedContent > iframe') ==False: # iframe 창이 아직 없으면 계속 눌러라
        try:
            driver.find_element_by_css_selector(f'#container > div > table.tb_base > tbody > tr:nth-child({num}) > td:nth-child(18) > a').click() # 1번 해설 업로드
            if check_exists_by_css_selector('#cboxLoadedContent > iframe') ==True:
                break
        except NoSuchElementException:
            pass
    driver.implicitly_wait(3)
    
    time.sleep(0.3)

    el = driver.find_element_by_class_name('cboxIframe') # iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
    driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨
    
    
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys('{0}'.format(list_sol_hml_files[idxnum])) # 해설 hml업로드
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload2"]').send_keys('{0}'.format(list_sol_png_files[idxnum])) # 해설 png업로드
    driver.find_element_by_css_selector('#numberOfAnswer_1').click() # 정답개수 체크 -> 일단 기본으로 1개 선택 / 새끼문제일때는 다시

    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload4"]').send_keys('{0}'.format(list_presol_hwp_files[idxnum])) # 정답 hwp업로드
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload3"]').send_keys('{0}'.format(list_presol_png_files[idxnum])) # 정답 png업로드
    select_presol_type(num) # 정답 종류 체크
    driver.find_element_by_css_selector('#txtAnswer').clear()
    driver.find_element_by_css_selector('#txtAnswer').send_keys('{0}'.format(presol_text[idxnum])) # 정답 텍스트 쓰기
    if list_presol_types[idxnum] == '[새끼문제]':
        select_son_cnt(num)
        driver.find_element_by_css_selector('#txtAnswer').clear()
    
    driver.find_element_by_css_selector('#btnSave').click()
    time.sleep(0.1)

    driver.switch_to.alert
    Alert(driver).accept()
    time.sleep(0.1)
    Alert(driver).accept()

    driver.switch_to.default_content() # 처음 frame으로 돌아가기
    time.sleep(0.1)


def select_presol_type(num):
    idxnum = num -1
    if list_presol_types[idxnum] == '[빈해설파일]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(1)').click()
    elif list_presol_types[idxnum] == '[객관식(선다-단일)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(2)').click()
    elif list_presol_types[idxnum] == '[객관식(선다-다중)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(3)').click()
    elif list_presol_types[idxnum] == '[객관식(OX-단일)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(4)').click()
    elif list_presol_types[idxnum] == '[객관식(OX-다중)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(5)').click()
    elif list_presol_types[idxnum] == '[주관식(정수)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(6)').click()
    elif list_presol_types[idxnum] == '[주관식(자판)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(9)').click()
    elif list_presol_types[idxnum] == '[증명문제]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(10)').click()
    elif list_presol_types[idxnum] == '[주관식(iink)]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(11)').click()
    elif list_presol_types[idxnum] == '[새끼문제]':
        driver.find_element_by_css_selector('#answerType > option:nth-child(12)').click()
    elif list_presol_types[idxnum] == '':
        # print("공통부분이어야 하는데 왜 나오지??")
        pass
    else:
        # print("엥 불가능한 경우 아닌가??")
        pass


def select_son_cnt(num):
    idxnum = num -1 
    renum = cnt_son_problems[idxnum]
    driver.find_element_by_css_selector('#numberOfAnswer_{0}'.format(renum)).click()




def dir_filelists_save(dir):
    global list_pro_hml_files, list_pro_png_files, list_sol_hml_files, list_sol_png_files, list_presol_hwp_files, \
        list_presol_types, list_presol_png_files, list_common_hml_files, list_common_png_files, list_common_first_last, \
            presol_text, cnt_son_problems
    list_pro_hml_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[1문제] [1hml]" in i] # dir 중 "[1문제] [1hml]" 있는 파일 제목들(경로포함)  
    list_pro_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[1문제] [2png]" in i] # dir 중 "[1문제] [2png]" 있는 파일 제목들(경로포함) 
    list_sol_hml_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[2해설] [1hml]" in i] # dir 중 "[1문제] [1hml]" 있는 파일 제목들(경로포함)  
    list_sol_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[2해설] [2png]" in i] # dir 중 "[2해설] [2png]" 있는 파일 제목들(경로포함) 
    list_presol_hwp_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[3정답] [1hwp]" in i] # dir 중 "[3정답] [1hwp]" 있는 파일 제목들(경로포함)
    list_presol_types = [i.split('.')[-2] for i in list_presol_hwp_files] # list_presol_hwp_files 중 정답 타입 부분만 
    list_presol_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[3정답] [2png]" in i] # dir 중 "[3정답] [2png]" 있는 파일 제목들(경로포함)
    list_common_hml_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[0공통문제] [1hml]" in i] # dir 중 "[0공통문제] [1hml]" 있는 파일 제목들(경로포함)  
    list_common_png_files = [os.path.join(dir, i) for i in os.listdir(dir) if "[0공통문제] [2png]" in i] # dir 중 "[0공통문제] [2png]" 있는 파일 제목들(경로포함)  
    list_common_first_last = [(i.split('번 [0공통문제]')[0][-7:].split('-')[0], i.split('번 [0공통문제]')[0][-7:].split('-')[1]) for i in list_common_hml_files] # [공통문제]중 (시작번호, 끝번호) 튜플형태로 추출


    presol_text = [] # list_presol_hwp_files 정답 파일 안의 텍스트들만 리스트화
    cnt_son_problems = []
    for list_presol_hwp_file in list_presol_hwp_files:
        f = olefile.OleFileIO(os.path.join(dir, list_presol_hwp_file)) # HWP 파일 열기
        encoded_text = f.openstream('PrvText').read() # PrvText 스트림의 내용 꺼내기
        decoded_text = encoded_text.decode('UTF-16') # 유니코드를 UTF-16으로 디코딩
        re_decoded_textf = re.sub('\n|\r', "", decoded_text)
        re_decoded_texts = re.sub('그리고', "$OTIMES$", re_decoded_textf)
        re_decoded_text = re.sub('또는', "$OPLUS$", re_decoded_texts)
        # print(re_decoded_text)
        presol_text.append(re_decoded_text)
        son_problems = re.findall('@', decoded_text)
        cnt_son_problems.append(len(son_problems))
        f.close()


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
            cnt_son_problems.insert(int(first)-1, "")







#### 팅거로 시작 부분 

# 일단 웹브라우져는 켜야해
driver_path = chromedriver_autoinstaller.install() # 버젼에 맞춰 자동 설치
options = webdriver.ChromeOptions()
options.add_argument("headless") # 창숨기는 옵션
options.add_argument("disable-gpu") # 불필요한 그래픽카드 기능을 제거함으로 셀레니움 작동속도 up
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path = driver_path, options=options)
driver.maximize_window() # 윈도우창 최대
driver.implicitly_wait(3) # 활성화 될때까지 최대 3초 기다려



root = Tk()
root.title("PM_divpro_uploader")


def browse_dest_path():
    folder_selected = filedialog.askdirectory()
    if folder_selected == "": # 사용자가 취소를 누를 때
        # print("폴더 선택 취소")
        return
    txt_dest_path.delete(0, END)
    txt_dest_path.insert(0, folder_selected.replace("/", "\\"))
    global dir
    dir = folder_selected.replace("/", "\\")


def start():
    # 파일 목록 확인
    if txt_dest_path.get() == '':
        msgbox.showwarning("경고", "업로드 작업 폴더를 선택하세요.")
        return
    elif pronum_head.get() == '':
        msgbox.showwarning("경고", "타이포서버 문제집 번호를 입력하세요.")
        return
    else:
        global typonum
        typonum = int(pronum_head.get())
        
        progress_condi.delete(1.0, END)
        progress_condi.update()

        dir_filelists_save(dir) # 작업폴더 각 파일을 리스트화해
        typo_open() # 타이포 열어서 몇번문제집까지 들어가 -> 500개 보기눌러놓고

        sol_up_parts = [] # 해설 업로드 파트 존재 여부 True, False로 받기
        for i in range(1, len(list_pro_hml_files)+1):
            driver.implicitly_wait(0.01)
            if check_exists_by_css_selector(f'#container > div > table.tb_base > tbody > tr:nth-child({i}) > td:nth-child(18) > a') == True:
                sol_up_parts.append('True')
            else :
                sol_up_parts.append('False')
            driver.implicitly_wait(3)

        for i in range(1, len(list_pro_hml_files)+1): # 문제 업로드 돌려
            pro_upload(i)
            progress_condi.insert(END, f"{i}번째 문제 업로드 완료\n")
            progress_condi.see(END)
            progress_condi.update()

        time.sleep(1)
        # print('문제 업로드 완료')

        for i in range(1, len(list_pro_hml_files)+1): # 해설 업로드 돌려 -> 돌리는 중 해설업로드 버튼이 없으면 넘어가 
            re_i = i-1
            if sol_up_parts[re_i]=='False':
                # print('넘어간다.')
                continue
            sol_upload(i)
            progress_condi.insert(END, f"{i}번째 해설 업로드 완료\n")
            progress_condi.see(END)
            progress_condi.update()

        time.sleep(1)


        time.sleep(1) # 정답 한글파일로 따로 받은 경우 -> 해설 디텍션은 안되어있는데 한글파일 정답은 있는 경우 존재 
        te = []
        for i in range(1, len(list_pro_hml_files)+1):
            if sol_up_parts[i-1]=='False' :
                if list_presol_types[i-1] == '[빈해설파일]':
                    print('Good')
                elif list_presol_types[i-1] == '':
                    print('공통문제임')
                else:
                    print('{0}번째 문제 잘못됨'.format(i))
                    te.append(i)
            else :
                pass

        if len(te) !=0 :
            # driver.find_element_by_css_selector('#txtTitle').send_keys('[해설디텍X, 해설파일O]') # 민약 위와 같은 문제가 있으면 제목에 체킹 
            msgbox.showinfo("알림", "[해설디텍]은 없는데 [해설파일]이 있는 파일입니다.\n \n!!!작업시트에 [해설디텍X, 해설파일O] 라고 체크해주세요.!!!")
        
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        msgbox.showinfo("알림", "작업이 완료되었습니다.")
        txt_dest_path.delete(1.0, END)
        pronum_head.delete(1.0, END)
        



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

pronum_head = Entry(pronum_frame)
pronum_head.pack(fill="both", padx=5, pady=5, ipady=4)






# 진행 상태 확인
progress_frame = LabelFrame(root, height=5, text="진행상황")
progress_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(progress_frame)
scrollbar.pack(side="right", fill="y")

progress_condi = Text(progress_frame, height=30, yscrollcommand=scrollbar.set)
progress_condi.pack(side="left", fill="both", expand=True)
scrollbar.config(command=progress_condi.yview)




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


driver.close()