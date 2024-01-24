import os.path
import sys
import time
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl

# ChromeOptions 및 WebDriverManager 설정
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("headless")
options.add_argument('--log-level=3')
options.add_argument('--disable-loging')
options.add_experimental_option("useAutomationExtension", False)
service = Service(executable_path=ChromeDriverManager().install())

answer업종 = ""
answer고용허가제 = ""
answer직무내용 = ""
answer모집인원 = ""
answer근무지 = ""
answer임금조건 = ""

# 기타 변수 정의
excelPath = "C:\jobdata\jobData.xlsx"
workbook = openpyxl.Workbook()

# 기존 엑셀 파일이 존재하는 경우 불러오기
if os.path.exists(excelPath):
    workbook = openpyxl.load_workbook(excelPath)

# 엑셀 시트 및 컬럼 설정
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
maxPagelen = 12

# 메인 링크 입력 받기
for curr in range(1, maxPagelen + 1):
    mainLink = ("https://www.work.go.kr/empInfo/empInfoSrch/list/dtlEmpSrchList.do?"
        "careerTo=&keywordJobCd=&occupation=&templateInfo=&shsyWorkSecd=&rot2WorkYn=&payGbn=&resultCnt=10&keywordJobCont=N"
        "&cert=&cloDateStdt=&moreCon=more&minPay=&codeDepth2Info=11000&isChkLocCall=&sortFieldInfo=DATE&major="
        "&resrDutyExcYn=&eodwYn=&sortField=DATE&staArea=&sortOrderBy=DESC&keyword=&termSearchGbn=all&carrEssYns="
        "&benefitSrchAndOr=O&disableEmpHopeGbn=&webIsOut=&actServExcYn=&maxPay=&keywordStaAreaNm=N&emailApplyYn="
        "&listCookieInfo=DTL&pageCode=&codeDepth1Info=11000&keywordEtcYn=&publDutyExcYn=&keywordJobCdSeqNo=&exJobsCd="
        "&templateDepthNmInfo=&computerPreferential=&regDateStdt=&employGbn=&empTpGbcd=&region=&infaYn=&resultCntInfo=10"
        f"&siteClcd=all&cloDateEndt=&sortOrderByInfo=DESC&currntPageNo={curr}&indArea=&careerTypes=&searchOn=Y&tlmgYn=&subEmpHopeYn="
        "&academicGbn=&templateDepthNoInfo=&foriegn=&mealOfferClcd=&station=&moerButtonYn=&holidayGbn=&srcKeyword="
        "&enterPriseGbn=all&academicGbnoEdu=noEdu&cloTermSearchGbn=all&keywordWantedTitle=N&stationNm=&benefitGbn="
        "&keywordFlag=&notSrcKeyword=&essCertChk=&isEmptyHeader=&depth2SelCode=&_csrf=dde2822b-952e-477d-81c3-17bb0c5d1775"
        "&keywordBusiNm=N&preferentialGbn=&rot3WorkYn=&pfMatterPreferential=&regDateEndt=&staAreaLineInfo1=11000"
        f"&staAreaLineInfo2=1&pageIndex={curr}&termContractMmcnt=&careerFrom=&laborHrShortYn=#viewSPL")
    
    # 메인 WebDriver 생성 및 메인 링크 접속
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(mainLink)

    # 메인 페이지에서 joblinks 수집
    joblinks = driver.find_elements(By.CLASS_NAME, 'cp-info-in')
    print(len(joblinks))

    maxPagelen = driver.find_element(By.CSS_SELECTOR, ".paging_direct").text.split(" ").pop(1)
    print(mainLink.split("pageIndex").pop(1))
    print(maxPagelen)
    
    # joblinks 반복 처리
    for linkss in joblinks:
        driverDetail = webdriver.Chrome(service=service, options=options)
        
        aTag = linkss.find_element(By.TAG_NAME, "a")
        aTagLink = aTag.get_attribute("href")

        driverDetail.get(aTagLink)
        
        try:
            elements_careers_table = set(driverDetail.find_elements(By.CLASS_NAME, 'careers-table'))
            elements_업종_right = driverDetail.find_element(By.CLASS_NAME, 'right')
            elements_업종_li = elements_업종_right.find_elements(By.TAG_NAME, "li")
            elements_업종_strong = elements_업종_right.find_element(By.CLASS_NAME, "info").find_elements(By.TAG_NAME, "strong")

            # 업종 정보 수집
            for idx, val in enumerate(elements_업종_strong):
                if val.text == "업종":
                    answer업종 = elements_업종_li.pop(idx).find_element(By.TAG_NAME, "div").text

            # 나머지 정보 수집
            for ele in elements_careers_table:
                tableTh = ele.find_elements(By.TAG_NAME, "th")

                for idx, val in enumerate(tableTh):
                    if val.text == "직무내용":
                        answers = ele.find_element(By.TAG_NAME, "td")
                        answer직무내용 = answers.text

                    if val.text == "모집인원":
                        answers = ele.find_elements(By.TAG_NAME, "td")
                        answer모집인원 = answers.pop(idx).text

                    if val.text == "근무예정지":
                        answers = ele.find_elements(By.TAG_NAME, "td")
                        answer근무지 = answers.pop(idx).text

                    if val.text == "임금조건":
                        answers = ele.find_elements(By.TAG_NAME, "td")
                        answer임금조건 = answers.pop(idx).text

                    if val.text == "고용허가제":
                        answers = ele.find_elements(By.TAG_NAME, "td")
                        answer고용허가제 = answers.pop(idx).text

                        # 고용허가제가 있을 경우 데이터 저장
                        if answer고용허가제 != " ":
                            assertResult = aTagLink
                            print(assertResult)
                            sheet.append([answer업종, answer직무내용, answer모집인원, answer근무지, answer임금조건, assertResult])
                            workbook.save("C:\jobdata\jobData.xlsx")
                            
        except NoSuchElementException as e:
            print(f"요소를 찾을 수 없습니다. 에러: {e.msg}")
            pass  # 요소를 찾을 수 없으면 패스              
        finally:
            driverDetail.close()
            
    driver.close()

# 엑셀 저장 및 WebDriver 종료
workbook.save("C:\jobdata\jobData.xlsx")
print("종료")

driverDetail.quit()
driver.quit()