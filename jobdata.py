from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import os.path
import sys
import time

# input = sys.stdin.readline

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("headless")
options.add_argument('--log-level=3') # 브라우저 로그 레벨을 낮춤
options.add_argument('--disable-loging') # 로그를 남기지 않음
options.add_experimental_option("useAutomationExtension", False)
service = Service(executable_path=ChromeDriverManager().install())
assertResult = ""
flag = False
mainLink = ""
answer업종 = ""
answer고용허가제 = ""
answer직무내용 = ""
answer모집인원 = ""
answer근무지 = ""
answer임금조건 = ""

workbook = Workbook()
sheet = workbook.active
sheet.column_dimensions['A'].width = 20
sheet.column_dimensions['B'].width = 50
sheet.column_dimensions['C'].width = 20
sheet.column_dimensions['D'].width = 20
sheet.column_dimensions['E'].width = 50


sheet.print_options.horizontalCentered = True
sheet.print_options.verticalCentered = True

sheet['A1'] = "업종"
sheet['B1'] = "직무내용"
sheet['C1'] = "모집인원"
sheet['D1'] = "근무지역"
sheet['E1'] = "임금조건"
sheet['F1'] = "URL 공고 주소"

mainLink = input(" 웹 주소 : ")

driver = webdriver.Chrome(service=service, options=options)
driver.get(mainLink)
# driver.implicitly_wait(120)

joblinks = driver.find_elements(By.CLASS_NAME, 'cp-info-in')
print(len(joblinks))

for linkss in joblinks:
    driverDetail = webdriver.Chrome(service=service, options=options)

    aTag = linkss.find_element(By.TAG_NAME, "a")
    aTagLink = aTag.get_attribute("href")
    
    driverDetail.get(aTagLink)
    # driverDetail.implicitly_wait(5)
    elements_careers_table = set(driverDetail.find_elements(By.CLASS_NAME, 'careers-table'))
    elements_업종_right = driverDetail.find_element(By.CLASS_NAME, 'right')
    elements_업종_li = elements_업종_right.find_elements(By.TAG_NAME, "li")
    elements_업종_strong = elements_업종_right.find_element(By.CLASS_NAME, "info").find_elements(By.TAG_NAME, "strong")
    
    for idx, val in enumerate(elements_업종_strong):
        if(val.text == "업종"):
            answer업종 = elements_업종_li.pop(idx).find_element(By.TAG_NAME, "div").text
    
    for ele in elements_careers_table:
        tableTh = ele.find_elements(By.TAG_NAME, "th")
        
        for idx, val in enumerate(tableTh):
            if(val.text == "직무내용"):
                answers = ele.find_element(By.TAG_NAME, "td")
                answer직무내용 = answers.text
                
            if(val.text == "모집인원"):
                answers = ele.find_elements(By.TAG_NAME, "td")
                answer모집인원 = answers.pop(idx).text                
                
            if(val.text == "근무예정지"):
                answers = ele.find_elements(By.TAG_NAME, "td")
                answer근무지 = answers.pop(idx).text
                
            if(val.text == "임금조건"):
                answers = ele.find_elements(By.TAG_NAME, "td")
                answer임금조건 = answers.pop(idx).text
                
            if(val.text == "고용허가제"):
                answers = ele.find_elements(By.TAG_NAME, "td")
                answer고용허가제 = answers.pop(idx).text
                
                if(answer고용허가제 != " "):
                    assertResult = aTagLink
                    print(assertResult)
                    sheet.append([answer업종, answer직무내용, answer모집인원, answer근무지, answer임금조건, assertResult])
                    
                    
    driverDetail.close()
    

workbook.save("C:\jobdata\jobData.xlsx")
print("end")
driver.close()

driverDetail.quit()
driver.quit()
