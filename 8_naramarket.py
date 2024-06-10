import os
from io import BytesIO
from time import sleep
from urllib.request import urlretrieve as download

import pandas as pd
import win32clipboard  # !pip install pywin32
import win32com.client as win32
from PIL import Image  # !pip install Pillow
from openpyxl import Workbook  # !pip install openpyxl
from selenium import webdriver  # !pip install selenium
from selenium.common.exceptions import *
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

def execute_script(script): # 자바스크립트 로딩 에러를 방지하기 위한 헬퍼함수
    while True:
        try:
            driver.execute_script(script)
            break
        except JavascriptException:
            sleep(0.5)

options = Options()
options.add_experimental_option('detach', True)  # 브라우저 바로 닫힘 방지
options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 불필요한 메시지 제거

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.implicitly_wait(5)
driver.get('http://shopping.g2b.go.kr/')
driver.switch_to.frame("sub")
search_input = driver.find_element(By.CSS_SELECTOR, 'input#kwd.srch_txt')
search_input.send_keys('작업용 의자')
search_input.submit()




# download_dir = r"C:/Users/user/Desktop/phthon/2_hg/market"
# driver = webdriver.Chrome(r"C:\Users\chromedriver.exe")'''