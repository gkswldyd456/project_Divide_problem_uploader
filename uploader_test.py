import time
import pyautogui
import os, shutil, sys
import fnmatch 

from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException

from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert

import olefile
import re




def typo_open():
    url = 'https://typo.postmath.co.kr/Web2/WorkbookList.aspx' # 타이포 서버
    driver.get(url) # 페이지 열어라 

    driver.find_element_by_css_selector('#mb_id').send_keys('postmath21') # 포메 아이디 입력
    driver.find_element_by_css_selector('#mb_password').send_keys('wD5V8Wo9@efHe$jW') # 포메 비번 입력

    if driver.find_element_by_xpath('//*[@id="id_save"]').get_attribute('checked') == None : # 아이디 저장 체크 되어있는지 판단 (안눌려있으면 None, 눌려있으면 True) 
        driver.find_element_by_xpath('//*[@id="id_save"]').click() # -> 안눌려 있다면 눌러라

    driver.find_element_by_css_selector('#btnLogin').click() # 로그인버튼 눌러
    time.sleep(0.2)

    typonum = int(pyautogui.prompt(text='문제집 번호(파일 번호)를 쓰시오.')) # 문제집번호
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




# for i in range(1, len(list_pro_hml_files)+1):
#     pro_upload(i)



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
        print("공통부분이어야 하는데 왜 나오지??")
    else:
        print("엥 불가능한 경우 아닌가??")



def select_son_cnt(num):
    idxnum = num -1 
    renum = cnt_son_problems[idxnum]
    driver.find_element_by_css_selector('#numberOfAnswer_{0}'.format(renum)).click()






###### 여기 부터 시작

dir = r"C:\Users\HanJiYong\Desktop\Testhaha\[1]하하하하하"

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
    re_decoded_text = re.sub('\n|\r', "", decoded_text)
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



driver_path = chromedriver_autoinstaller.install() # 버젼에 맞춰 자동 설치
options = webdriver.ChromeOptions()
# options.add_argument("headless") # 창숨기는 옵션
# options.add_experimental_option("detach", True) # 뭐 분리해서 안꺼지게 해준다는데 잘 안됨....ㅜㅜ
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path = driver_path, options=options)
driver.maximize_window() # 윈도우창 최대
driver.implicitly_wait(3) # 활성화 될때까지 최대 3초 기다려



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

time.sleep(1)
print('문제 업로드 완료')

for i in range(1, len(list_pro_hml_files)+1): # 해설 업로드 돌려 -> 돌리는 중 해설업로드 버튼이 없으면 넘어가 
    re_i = i-1
    if sol_up_parts[re_i]=='False':
        print('넘어간다.')
        continue
    sol_upload(i)

time.sleep(1)
print('해설 업로드 완료')





# time.sleep(1) # 정답 한글파일로 따로 받은 경우 -> 해설 디텍션은 안되어있는데 한글파일 정답은 있는 경우 존재 
# te = []
# for i in range(1, len(list_pro_hml_files)+1):
#     if sol_up_parts[i-1]=='False' :
#         if list_presol_types[i-1] == '[빈해설파일]':
#             print('Good')
#         elif list_presol_types[i-1] == '':
#             print('공통문제임')
#         else:
#             print('{0}번째 문제 잘못됨'.format(i))
#             te.append(i)
#     else :
#         pass

# if len(te) !=0 :
#     driver.find_element_by_css_selector('#txtTitle').send_keys('[해설디텍X, 해설파일O]') # 민약 위와 같은 문제가 있으면 제목에 체킹 




















# ### 1회째

# while check_exists_by_css_selector('#cboxLoadedContent > iframe') ==False: # iframe 창이 아직 없으면 계속 눌러라
#     try:
#         driver.find_element_by_css_selector('#container > div > table.tb_base > tbody > tr:nth-child(1) > td:nth-child(17) > a').click() # 1번 문제 업로드
#         if check_exists_by_css_selector('#cboxLoadedContent > iframe') ==True:
#             break
#     except NoSuchElementException:
#         pass

# el = driver.find_element_by_class_name('cboxIframe') # iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
# driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨
# time.sleep(1)

# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys('{0}'.format(list_pro_hml_files[1])) # 문제 hml업로드
# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload2"]').send_keys('{0}'.format(list_pro_png_files[1])) # 문제 png업로드
# driver.find_element_by_css_selector('#btnSave').click()
# time.sleep(0.1)

# driver.switch_to.alert
# Alert(driver).accept()
# time.sleep(0.1)
# Alert(driver).accept()

# driver.switch_to.default_content() # 처음 frame으로 돌아가기
# time.sleep(0.1)

# ### 2회째

# while check_exists_by_css_selector('#cboxLoadedContent > iframe') ==False: # iframe 창이 아직 없으면 계속 눌러라
#     try:
#         driver.find_element_by_css_selector('#container > div > table.tb_base > tbody > tr:nth-child(2) > td:nth-child(17) > a').click() # 1번 문제 업로드
#         if check_exists_by_css_selector('#cboxLoadedContent > iframe') ==True:
#             break
#     except NoSuchElementException:
#         pass

# el = driver.find_element_by_class_name('cboxIframe') # iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
# driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨
# # time.sleep(1)

# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys('{0}'.format(list_pro_hml_files[2])) # 문제 hml업로드
# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload2"]').send_keys('{0}'.format(list_pro_png_files[2])) # 문제 png업로드
# driver.find_element_by_css_selector('#btnSave').click()
# time.sleep(0.1)

# driver.switch_to.alert
# Alert(driver).accept()
# time.sleep(0.1)
# Alert(driver).accept()

# driver.switch_to.default_content() # 처음 frame으로 돌아가기
# time.sleep(0.1)



# driver.find_element_by_css_selector('#container > div > table.tb_base > tbody > tr:nth-child(2) > td:nth-child(17) > a').click() # 2번 문제 업로드

# el = driver.find_element_by_class_name('cboxIframe') # pdf업로드 화면이 iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
# driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨
# # time.sleep(1)

# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys(r"C:\Users\HanJiYong\Desktop\Testhaha\[1]하하하하\[1]하하하하 002번 [1문제] [1hml].hml")
# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload2"]').send_keys(r"C:\Users\HanJiYong\Desktop\Testhaha\[1]하하하하\[1]하하하하 002번 [1문제] [2png].png")
# driver.find_element_by_css_selector('#btnSave').click()

# driver.switch_to.default_content() # 처음 frame으로 돌아가기




#btnCancel
#btnSave
# driver.find_element_by_css_selector('#container > div > table.tb_base > tbody > tr:nth-child(1) > td:nth-child(17) > a').click() # tr:nth-child(1) > td:nth-child(17) -> 1번 >  문제 업로드



#container > div > table.tb_base > tbody > tr:nth-child(1) > td:nth-child(17) > a





# tabs = driver.window_handles
# driver.execute_script('window.open("https://typo.postmath.co.kr/Web2/WorkbookDetailList.aspx?seq=1101");')
# time.sleep(0.2)

#container > div > table.tb_base > tbody > tr:nth-child(120) > td:nth-child(17) > a  # tr:nth-child(120) > td:nth-child(17) -> 120번 문제 업로드
#container > div > table.tb_base > tbody > tr:nth-child(120) > td:nth-child(18) > a  # tr:nth-child(120) > td:nth-child(18) -> 120번 해설 업로드
# driver.get(re_url)

# driver.switch_to.window(driver.window_handles[0])
# time.sleep(0.2)
# driver.switch_to.window(driver.window_handles[1])
# time.sleep(0.2)
# driver.switch_to.window(driver.window_handles[0])
# time.sleep(0.2)
# driver.switch_to.window(driver.window_handles[1])
# time.sleep(0.2)
# driver.switch_to.window(driver.window_handles[0])
# time.sleep(0.2)
# # driver.get(re_url)

# driver.find_element_by_css_selector('#ContentPlaceHolder1_btnReg').click() # pdf업로드 버튼눌러
# time.sleep(1)

# # # 현재 웹페이지에서 iframe이 몇개가 있는지 변수에 넣고 확인해 봅니다.
# # iframes = driver.find_elements_by_tag_name('iframe')
# # print('현재 페이지에 iframe은 %d개가 있습니다.' % len(iframes))

# el = driver.find_element_by_class_name('cboxIframe') # pdf업로드 화면이 iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
# driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨
# time.sleep(1)
# #ddlSource11
# #ddlSource11 > option
# #ddlSource1 > option
# def FC():
#     html = driver.page_source
#     soup = BeautifulSoup(html)
#     Item_year = soup.select('#ddlSource11 > option') # 년도 안 정보
#     Item_source1 = soup.select('#ddlSource1 > option') # 출처1 안 정보
#     Item_source2 = soup.select('#ddlSource2 > option') # 출처2 안 정보
#     Item_source3 = soup.select('#ddlSource3 > option') # 출처3 안 정보
#     Item_source4 = soup.select('#ddlSource4 > option') # 출처4 안 정보
#     Item_source5 = soup.select('#ddlSource5 > option') # 출처5 안 정보
    
#     global item_years, item_source1, item_source2, item_source3, item_source4, item_source5
#     # item_years_value = [i["value"] for i in Item_year] # 만약 Item_year 정보 중 value 값을 가져오고 싶을 때
#     item_years = [i.text for i in Item_year] # 년도 텍스트만
#     item_source1 = [i.text for i in Item_source1]
#     item_source2 = [i.text for i in Item_source2]
#     item_source3 = [i.text for i in Item_source3]
#     item_source4 = [i.text for i in Item_source4]
#     item_source5 = [i.text for i in Item_source5]
    
    
    
# FC()
# print(item_years)
# print(item_source1)
# print(item_source2)
# print(item_source3)
# print(item_source4)
# print(item_source5)


# # print(item_years.index('2004년')) # 2004년에 대한 idx값 -> int형 


# driver.find_element_by_css_selector('#txtWorkbookTitle').send_keys('하하하하') # [문제집 제목]에 하하하하

# driver.find_element_by_css_selector('#ddlSource11 > option:nth-child({0})'.format(item_years.index('2004년')+1)).click() # 2004년에 해당하는 것 눌러

# driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys(r"C:\Users\HanJiYong\Desktop\하하하하 008번 [3정답] [3hwp].pdf")

# #ContentPlaceHolder1_fileUpload1

# time.sleep(3)


# driver.find_element_by_css_selector('#btnCancel').click() # 취소버튼
# time.sleep(1)

# driver.switch_to.default_content() # 처음 frame으로 돌아가기
# driver.find_element_by_css_selector('#ContentPlaceHolder1_btnReg').click() # pdf업로드 버튼눌러






### 정보TIP (검색해보면 좋음)
# os.path.split -> 폴더, 파일명 나눠주는거 튜플로 나눠줌
# os.path.join -> 경로 (또는 파일명) 합쳐줌
# os.listdir(dir) -> dir(폴더) 안에 파일들 검색해줘 (list로 저장될거임)
# os.remove -> 파일제거해
# os.rename - > 파일명 변경
# os.makedirs -> 폴더 만들어
# shutil.move -> 폴더 옮겨 
# fnmatch -> 이름매칭? 




# try:
#     driver.find_element_by_id('ContentPlaceHolder1_fileUpload1').is_enabled==True
#     print('파일선택 있다!! 누른다???')
#     driver.find_element_by_id('ContentPlaceHolder1_fileUpload1').click()
# except NoSuchElementException:
#     print('파일선택 없다... ㅠㅠㅜㅠ')



# 작업 끝나고 구글창 안꺼지게 무한 루프
while True: # 
    pass



