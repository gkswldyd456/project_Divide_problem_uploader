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


driver_path = chromedriver_autoinstaller.install() # 버젼에 맞춰 자동 설치
options = webdriver.ChromeOptions()
# options.add_argument("headless") # 창숨기는 옵션
# options.add_experimental_option("detach", True) # 뭐 분리해서 안꺼지게 해준다는데 잘 안됨....ㅜㅜ
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path = driver_path, options=options)
driver.maximize_window() # 윈도우창 최대
driver.implicitly_wait(3) # 활성화 될때까지 최대 3초 기다려

url = 'https://typo1.postmath.co.kr/Web2/WorkbookList.aspx' # 타이포 1

# url = 'https://typo.postmath.co.kr/Web2/WorkbookList.aspx' # 타이포 서버
driver.get(url) # 페이지 열어라 

driver.find_element_by_css_selector('#mb_id').send_keys('postmath21') # 포메 아이디 입력
driver.find_element_by_css_selector('#mb_password').send_keys('wD5V8Wo9@efHe$jW') # 포메 비번 입력

if driver.find_element_by_xpath('//*[@id="id_save"]').get_attribute('checked') == None : # 아이디 저장 체크 되어있는지 판단 (안눌려있으면 None, 눌려있으면 True) 
    driver.find_element_by_xpath('//*[@id="id_save"]').click() # -> 안눌려 있다면 눌러라

driver.find_element_by_css_selector('#btnLogin').click() # 로그인버튼 눌러
time.sleep(0.2)





dir = r"C:\Users\HanJiYong\Desktop\네이밍" # 파일 있는 경로명
moi_files_dirOK = [os.path.join(dir, i) for i in os.listdir(dir)] # 모의고사 파일 리스트화 (경로포함)
moi_files = [i.replace(".pdf", "") for i in os.listdir(dir)]





def pdf_upload(num):
    idxnum = num -1 
    driver.find_element_by_css_selector('#ContentPlaceHolder1_btnReg').click() # pdf업로드 버튼눌러
    time.sleep(1)

    el = driver.find_element_by_class_name('cboxIframe') # pdf업로드 화면이 iframe(웹사이트 안에 웹사이트를 부른거라 생각하면됨)
    driver.switch_to.frame(el) # 그 iframe으로 포커스 옮겨

    driver.find_element_by_css_selector('#txtWorkbookTitle').send_keys('{0}'.format(moi_files[idxnum])) # 문제집 제목 넣기
    driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_fileUpload1"]').send_keys('{0}'.format(moi_files_dirOK[idxnum])) # 파일 넣기
    time.sleep(0.1)
    driver.find_element_by_css_selector('#btnSave').click()
    time.sleep(0.1)
    
    # driver.switch_to.alert
    # Alert(driver).accept()
    # time.sleep(0.1)
    # Alert(driver).accept()

    # driver.switch_to.default_content() # 처음 frame으로 돌아가기
    # time.sleep(0.1)

pdf_upload(1)


# 작업 끝나고 구글창 안꺼지게 무한 루프
while True: # 
    pass