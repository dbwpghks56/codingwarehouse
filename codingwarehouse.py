import logging
import os.path
import pickle
import sys
import time
from urllib.parse import parse_qs, urlparse

import openpyxl
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# url param 추출
def extract_param_value(url):
    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    
    if "wantedAuthNo" in query_params:
        return query_params["wantedAuthNo"][0]
    else:
        return None

# 로깅 설정
logging.basicConfig(filename='my_log_file.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# login driver 설정
optionsLogin = webdriver.ChromeOptions()
optionsLogin.add_experimental_option('excludeSwitches', ['enable-logging'])
optionsLogin.add_experimental_option("excludeSwitches", ["enable-automation"])
optionsLogin.add_experimental_option("useAutomationExtension", False)

# ChromeOptions 및 WebDriverManager 설정
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("headless")
options.add_argument('--log-level=3')
options.add_argument('--disable-loging')
options.add_experimental_option("useAutomationExtension", False)
service = Service(executable_path=ChromeDriverManager().install())

answer문제번호 = ""
answer문제제목 = ""
answer문제내용 = ""
answer시간제한 = ""
answer메모리제한 = ""
answer총시도 = 0
answer맞춘시도 = 0
answer맞힌사람 = 0
answer정답비율 = ""
answer난이도 = ""
answer태그 = ""
answer출처 = ""
saveFlag = False
stopFlag = False

cookiefilelink = "C:\codingwarehouse\\baekjoonlogin.pkl"
firstfilelink = "C:\codingwarehouse\\firstInfo.txt"
# 기타 변수 정의
excelPath = "C:\codingwarehouse\\baekjoon.xlsx"

workbook = openpyxl.Workbook()

# 기존 엑셀 파일이 존재하는 경우 불러오기
if os.path.exists(excelPath):
    workbook = openpyxl.load_workbook(excelPath)


# 엑셀 시트 및 컬럼 설정
sheet = workbook.active
sheet.print_options.horizontalCentered = True
sheet.print_options.verticalCentered = True
sheet['A1'] = "문제번호"
sheet['B1'] = "문제제목"
sheet['C1'] = "문제내용"
sheet['D1'] = "시간제한"
sheet['E1'] = "메모리제한"
sheet['F1'] = "총시도"
sheet['G1'] = "맞춘시도"
sheet['H1'] = "맞힌사람"
sheet['I1'] = "정답비율"
sheet['J1'] = "난이도"
sheet['k1'] = "태그"
sheet['l1'] = "출처"

maxPagelen = 189


currPage = input("검색 시작할 페이지 : ")
loginRedirectLink = ("https://www.acmicpc.net/")
loginMainLink = ("https://www.acmicpc.net/login?next=%2F")
mainLink = (f"https://www.acmicpc.net/problemset?sort=solvedac_asc&tier=1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20&page={currPage}")

loginDriver = webdriver.Chrome(service=service, options=optionsLogin)

try:
    loginDriver.get(loginMainLink)
    wait = WebDriverWait(loginDriver, 120)
    
    def check_url(loginDriver):
        return loginRedirectLink in loginDriver.current_url and loginMainLink not in loginDriver.current_url
        
        # 조건 함수를 사용하여 기다리기
    wait.until(check_url)
    
    print(loginDriver.current_url)
    
    if loginRedirectLink in loginDriver.current_url and loginDriver.current_url != loginMainLink :
        print("hello")
        pickle.dump(loginDriver.get_cookies(), open(cookiefilelink, "wb"))
            
except Exception as e:
    print(e)
    pass

finally:
    pickle.dump(loginDriver.get_cookies(), open(cookiefilelink, "wb"))
    loginDriver.close()


# 메인 링크 입력 받기
for curr in range(int(currPage)+1, int(maxPagelen) + 1):
    try:
        # 메인 WebDriver 생성 및 메인 링크 접속
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(mainLink)

        # 메인 페이지에서 joblinks 수집
        tableId = driver.find_element(By.ID, 'problemset')
        tbody = tableId.find_element(By.TAG_NAME, 'tbody')
        tr = tbody.find_elements(By.TAG_NAME, 'tr')
        
        for idx, val in enumerate(tr):
            td = val.find_elements(By.TAG_NAME, 'td')
            aTag = td[1].find_element(By.TAG_NAME, 'a')
            aTagLink = aTag.get_attribute("href")
            
            driverDetail = webdriver.Chrome(service=service, options=options)
            
            if os.path.exists(cookiefilelink):
                workCookies = pickle.load(open(cookiefilelink, "rb"))
                driverDetail.get(loginRedirectLink)
                driverDetail.delete_all_cookies()
                
                for cookie in workCookies:
                    # cookie.pop("domain")
                    driverDetail.add_cookie(cookie)

                driverDetail.get(aTagLink)
            
            try:
                answer문제번호 = aTagLink.split("/")[-1]
                answer문제제목 = driverDetail.find_element(By.ID, 'problem_title').text
                dataTable = driverDetail.find_element(By.ID, 'problem-info')
                dataBody = dataTable.find_element(By.TAG_NAME, 'tbody')
                dataTr = dataBody.find_elements(By.TAG_NAME, 'tr')
                dataTd = dataTr[0].find_elements(By.TAG_NAME, 'td')
                
                answer시간제한 = dataTd[0].text
                answer메모리제한 = dataTd[1].text
                answer총시도 = dataTd[2].text
                answer맞춘시도 = dataTd[3].text
                answer맞힌사람 = dataTd[4].text
                answer정답비율 = dataTd[5].text
                
                answer문제내용 = driverDetail.find_element(By.ID, 'problem_description').text
                
                try:
                    clickA = driverDetail.find_element(By.CLASS_NAME, 'show-spoiler')
                    
                    clickA.click()
                except:
                    pass
                
                
                tagUl = driverDetail.find_elements(By.CLASS_NAME, 'spoiler-link')
                
                try:
                    source_ul = driverDetail.find_element(By.ID, 'source')
                    
                    answer출처 = source_ul.text
                except:
                    print("출처가 없습니다.")
                    pass
                
                for tag in tagUl:
                    answer태그 += tag.text + ", "
                
                sheet.append([answer문제번호, answer문제제목, answer문제내용, answer시간제한, answer메모리제한,
                                        answer총시도, answer맞춘시도, answer맞힌사람, answer정답비율, "", answer태그, answer출처])
                
                workbook.save(excelPath)
                
                answer태그 = ""
                
            except NoSuchElementException as e:
                print(f"요소를 찾을 수 없습니다. 에러: {e.msg}")
                pass  # 요소를 찾을 수 없으면 패스 
            except Exception as e2:
                print(f"에러 요인 {e2}")
                pass             
            finally:
                stopFlag = False
                driverDetail.close()
        
        mainLink = mainLink.replace(f"page={curr-1}", f"page={curr}")
        driver.close()
        
    except Exception as e:
        print(f"에러 요인 {e}")
        pass
    
    finally:
        stopFlag = False
        time.sleep(1)
        print(f"현재 페이지 {curr-1} 입니다.")
        print(f"최대 페이지 {maxPagelen} 입니다.")
        print("다음 페이지로 이동합니다.")
    
    # if stopFlag:
    #     break


# 엑셀 저장 및 WebDriver 종료
workbook.save(excelPath)

print("종료")

driverDetail.quit()
loginDriver.quit()
driver.quit()