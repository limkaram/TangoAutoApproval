##
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
import re
import pyautogui
import pyperclip

## headless 설정
# options = webdriver.ChromeOptions()
# options.add_argument('headless')
# options.add_argument('window-size=1920x1080')
# options.add_argument("disable-gpu")

##
print('!알림 : 외부망 승인 가능 계정 로그인')
userId = input('아이디 : ')
userPwd = input('비밀번호 : ')
# userId = 'SKN20001243'
# userPwd = 'Rkfka.411l@'

## TANGO 로그인 페이지 접속
# Chrome의 경우 아까 받은 chromedriver의 위치를 지정해준다.
login_url = 'https://www.tango.sktelecom.com/tango-common-business-web/login/Login.do'  # TANGO 로그인 페이지
# driver = webdriver.Chrome('C:\\Users\\SKNS\\PycharmProjects\\chromedriver', options=options)  # headless 활용
driver = webdriver.Chrome()  # 드라이버 접근
driver.get(login_url)  # TANGO 로그인 페이지 접속

## 아이디, 비밀번호 입력 후 로그인 버튼 클릭
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userId'))).send_keys(userId)  # 아이디 입력
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#userPwd'))).send_keys(userPwd)  # 비밀번호 입력
driver.find_element_by_xpath('/html/body/div[1]/div[1]/button').click()

##
time.sleep(1)
login_success = len(driver.window_handles) > 1  # 로그인 성공의 경우(인증번호 입력 팝업창 띄움)

## 인증번호를 통한 로그인
if login_success:  # 로그인 성공
    print('!알림 : 로그인 성공')
    driver.switch_to.window(driver.window_handles[1])  # 로그인 팝업창으로 스위칭
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#btnGetNumber'))).click()  # 인증번호 받기 클릭
    print('!알림 : 인증번호 발송 완료')
    authNumber = input('인증번호 : ')
    driver.find_element_by_css_selector('#authNumber').send_keys(authNumber)  # 인증번호 입력
    driver.find_element_by_css_selector('#btnOk').click()
    time.sleep(2)
    driver.find_element_by_xpath('/html/body/div[3]/div[3]/button').click()
else:  # 로그인 실패
    print('!알림 : 로그인 실패')
    print('!알림 : 아이디 및 비밀번호를 확인해주세요.')
    print('10초 후 프로그램이 종료됩니다.')
    time.sleep(10)
    driver.close()

##
driver.switch_to.window(driver.window_handles[0])  # 메인 윈도우로 변환

try:  # 최초 창 다시보지않음 선택한 유저는 없을 수 있으므로 패스시키고, 존재하면 클릭하여 지나감
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[1]/div[1]/a/img').click()
except:
    pass

## 공지사항 제거
try:  # 최초 창 다시보지않음 선택한 유저는 없을 수 있으므로 패스시키고, 존재하면 클릭하여 지나감
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[7]/div/button').click()
except:
    pass

##  운용-작업관리-To-Do List 진입
print('!알림 : To-Do List 진입 중...')
time.sleep(5)
driver.find_element_by_xpath('/html/body/header/div/div[3]/div[1]/button').click()  # 업무구분 클릭

for i in range(1, 6):
    try:
        tap = driver.find_element_by_xpath('/html/body/header/div/div[3]/div[1]/ul/li[{0}]/a'.format(i)).text
        print(tap)
        if tap == '운용':
            driver.find_element_by_xpath('/html/body/header/div/div[3]/div[1]/ul/li[{0}]/a'.format(i)).click()
    except:
        break

## 공지사항 팝업 존재의 경우
try:
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[7]/div/button').click()
except:
    pass

##
time.sleep(5)

try:
    driver.find_element_by_xpath('/html/body/header/div/div[3]/div[2]/ul/li[4]/a').click()  # 작업관리 클릭
except:
    driver.find_element_by_xpath('/html/body/header/div/div[3]/div[2]/ul/li[3]/a').click()  # 작업관리 클릭
driver.find_element_by_xpath('/html/body/header/div/div[1]/div/div[2]/ul[2]/li[2]/ul/li[1]/div').click()  # To-Do list 클릭
print('!알림 : To-Do List 진입 완료')

## 도메인 외부망만 체크
time.sleep(5)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div[1]/div[1]/label/div/label[1]').click()  # 전체 체크박스 클릭
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div[1]/div[1]/label/div/label[7]').click()  # 외부망 체크박스 클릭

##
today = datetime.today()
start_date_obj = datetime(year=today.year, month=today.month, day=today.day)
end_date_obj = start_date_obj + timedelta(days=7)  # 현재 local 시간 기준 7일 이후 날짜 설정
end_date_str = end_date_obj.strftime('%Y-%m-%d')

## 진행단계 : 작업검토+작업승인(팀장) 클릭. 반려 작업 제외 목적
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div[2]/div[1]/label/div[1]/select/option[4]').click()
time.sleep(1)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div[2]/div[1]/label/div[2]/select/option[3]').click()

## end date를 금일로부터 7일 뒤의 날짜로 넣고 외부망 작업 조회
end_date_click = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/div[2]/div[3]/div/div[2]/input')  # To 기간 클릭
for _ in range(10):
    end_date_click.send_keys(Keys.BACK_SPACE)
end_date_click.send_keys(end_date_str)
driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/form/div/button[2]').click()  # 조회 클릭
time.sleep(5)

## display 되는 작업의 갯수 늘리기(100개)
try:
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[2]/div[5]/div[3]/select/option[5]').click()
except:
    print('!알림 : 등재된 외부망 작업이 없습니다.')
    print('!알림 : 10초 뒤 프로그램이 종료됩니다.')
    time.sleep(10)

## 보안 뚫기
print('!알림 : TANGO 보안 뚫는 중...')
for i in range(1000):
    try:
        secret_xpath_obj = driver.find_element_by_xpath('//*[@id="alopexgrid{0}-0-6"]'.format(i))
        if secret_xpath_obj:
            secret_xpath_str = '//*[@id="alopexgrid{0}-0-6"]'.format(i)
            break
    except:
        pass

def xpath_job_by_idx(row, column, xpath):
    # 작업명 : //*[@id="alopexgrid374-0-6"]
    # 작업일시 : //*[@id="alopexgrid374-0-10"]
    # 등록일시 : //*[@id="alopexgrid374-0-11"]
    xpath_split = xpath.split('-')
    xpath_split[-2] = str(row)  # 작업명에 해당
    xpath_split[-1] = str(column) + '"]'
    result = '-'.join(xpath_split)
    return result
print('!알림 : TANGO 보안 ID 추출 완료')

##
def confirm_skt_time_rule(row):
    # 24시간 이후 등재 작업인지 확인
    worktime_colums_idx = 10  # 작업일시 컬럼 인덱스
    registtime_colums_idx = 11  # 등록일시 컬럼 인덱스

    work_time = driver.find_element_by_xpath(xpath_job_by_idx(row, worktime_colums_idx, secret_xpath_str)).text  # 작업일시
    work_month = int(work_time.split()[0].split('/')[0])
    work_day = int(work_time.split()[0].split('/')[1])
    work_datetime = datetime(year=today.year, month=work_month, day=work_day)

    regist_time = driver.find_element_by_xpath(xpath_job_by_idx(row, registtime_colums_idx, secret_xpath_str)).text  # 등록일시
    regist_month = int(regist_time.split()[0].split('-')[0])
    regist_day = int(regist_time.split()[0].split('-')[1])
    regist_datetime = datetime(year=today.year, month=regist_month, day=regist_day)

    next_regist_month = (regist_month + 1) % 12
    if next_regist_month <= work_month:  # 내년으로 변경되는 경우
        work_datetime = datetime(year=today.year+1, month=work_month, day=work_day)

    skt_regist_rule = work_datetime >= (regist_datetime + timedelta(days=1))

    return skt_regist_rule

## 위쪽 텍스트 정보 가져오기
def get_above_info():
    above_info = driver.find_element_by_xpath('/html').text.split('\n')

    above_info = [i for i in above_info
                if i.startswith('작업명') or
                i.startswith('작업예상일시') or
                i.startswith('작업 영향')]

    above_jobname = above_info[0]  # 위쪽 작업명 가져오기
    above_jobtime = above_info[1]  # 위쪽 작업예상일시 가져오기
    above_jobeffect = above_info[2]  # 위쪽 작업영향 가져오기
    # print(above_info)
    # print(above_jobname)
    # print(above_jobtime)
    # print(above_jobeffect)

    return above_jobname, above_jobtime, above_jobeffect

## 아래쪽 정보 가져오기
def get_under_info():
    # last_height = driver.execute_script("return document.body.scrollHeight")  # 스크롤 높이를 받아옴
    driver.execute_script("window.scrollTo(0, 470)")  # 470 위치까지 스크롤 내림

    # 마우스 위치 가져오기
    # x, y = pyautogui.position()
    # print('x={0}, y={1}'.format(x, y))

    # 드래그 시행
    # 시작점 : x=1438, y=685
    # 끝점 : x=667, y=261
    pyautogui.moveTo(1417, 680, 0.2)  # 0.5초 안에 해당 좌표로 이동
    pyautogui.mouseDown(x=1417, y=680)  # 이동한 좌표에서 클릭

    pyautogui.moveTo(667, 261, 0.2)  # 0.5초 안에 해당 좌표로 이동
    pyautogui.mouseUp(x=667, y=261)  # 이동한 좌표에서 클릭 종료

    pyautogui.hotkey('ctrl', 'c')
    text_copied = pyperclip.paste()

    under_info = text_copied.split('\r\n')

    for text in under_info:
        temp = text.replace(' ', '')  # 띄워쓰기 제거 후 작업명, 작업사유, 작업 일시 찾기

        if re.search(r'[작][ ]{0,1}[업][ ]{0,1}[명]', temp):
            under_jobname = text

        if re.search(r'[작][ ]{0,1}[업][ ]{0,1}[일][ ]{0,1}[시]', temp):
            under_jobtime = text

        if re.search(r'[출][ ]{0,1}[입][ ]{0,1}[국][ ]{0,1}[소]', temp):
            under_jobenterance = text

        if re.search(r'[서][ ]{0,1}[비][ ]{0,1}[스][ ]{0,1}[영][ ]{0,1}[향]', temp):
            under_jobeffect = text

    # under_jobname = under_info[0]
    # under_jobtime = under_info[1]
    # under_jobenterance = under_info[2]
    # under_jobeffect = under_info[3]
    # print(under_info)
    # print(under_jobname)
    # print(under_jobtime)
    # print(under_jobenterance)
    # print(under_jobeffect)

    return under_jobname, under_jobtime, under_jobenterance, under_jobeffect

##
def cleaning_the_jobname(name):
    preprocessed_names = []

    regexed_name = re.findall(r'[S]{1}[가-힣]{1,9}', name)
    if regexed_name is not None:
        preprocessed_names.extend(regexed_name)

    regexed_name = re.findall(r'[가-힣]{1,9}[_ ]{0,1}[전중집국통][송심중사소합][실국]{0,1}', name)
    if regexed_name is not None:
        preprocessed_names.extend(regexed_name)

    regexed_name = re.findall(r'[sS][kK][tTbB][_ ]{0,1}[가-힣]{1,9}', name)
    if regexed_name is not None:
        preprocessed_names.extend(regexed_name)

    # 예외 추가 사항 적용
    regexed_name = re.findall(r'[반][포][T][2]', name)
    if regexed_name is not None:
        preprocessed_names.extend(regexed_name)

    regexed_name = re.findall(r'[T][2][반][포]', name)
    if regexed_name is not None:
        preprocessed_names.extend(regexed_name)

    regexed_name = re.findall(r'[S][I][T][C]', name)
    if regexed_name is not None:
        preprocessed_names.extend(regexed_name)

    for preprocessed_name in preprocessed_names:
        preprocessed_name.replace(' ', '')  # 띄워쓰기 제거
        preprocessed_name.replace('_', '')  # 언더바 제거
        preprocessed_name.replace('S', '')  # 수도권 접두어 제거
        preprocessed_name.replace('SKT', '')  # 접두어 제거
        preprocessed_name.replace('SKB', '')  # 접두어 제거
        preprocessed_name.replace('skt', '')  # 접두어 제거
        preprocessed_name.replace('skb', '')  # 접두어 제거

    return preprocessed_names

## 승인 진행 영역
print('!알림 : 작업 검토 및 승인 시작')
row = 0
jobname_column_idx = 6
# 작업명(column:6), 작업일시(column:10), 등록일시(column:11)

while True:
    try:
        xpath = xpath_job_by_idx(row, jobname_column_idx, secret_xpath_str)  # 작업명을 클릭하는 것으로 정함
        job_name = driver.find_element_by_xpath(xpath_job_by_idx(row, jobname_column_idx, secret_xpath_str)).text
        print('============================================================================')
        print(' - 작업명 :', job_name)

        # 24시간 등재 기준 적합성 판단
        if confirm_skt_time_rule(row):  # 등재 후 24시간 이후 시행 작업
            pass
        else:  # 등재 후 24시간 이전 시행 작업
            if '긴급' in job_name:
                pass
            else:
                print(' - 결과 : 미승인')
                print(' - 사유 : 24시간 이내 등재 후 시행작업 긴급 타이틀 부재')
                row += 1
                continue

        # 외부망 작업 클릭 후 팝업창으로 전환
        driver.find_element_by_xpath(xpath).click()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[1])

        # 웹페이지내 웹페이지(iframe)으로 이동
        internal_webpage = driver.find_element_by_tag_name('iframe')  # 내부 웹페이지 가져오기
        driver.switch_to.frame(internal_webpage)  # 내부 웹페이지로 스위치

        # 팝업창내 위쪽 정보 가져오기
        above_jobname, above_jobtime, above_jobeffect = get_above_info()

        # 팝업창내 아래쪽 정보 가져오기
        under_jobname, under_jobtime, under_jobenterance, under_jobeffect = get_under_info()

        # 가져온 팝업내 정보 정제
        cleaned_jobname_above = cleaning_the_jobname(above_jobname)  # 상단 작업명
        cleaned_jobname_under = cleaning_the_jobname(under_jobname)  # 하단 작업명
        cleaned_jobname_jobenterance = cleaning_the_jobname(under_jobenterance)  # 하단 출입국소

        # iframe에서 정보를 모두 얻었으므로, 기존으로 다시 돌아옴
        driver.switch_to.default_content()

        # 서비스 영향 검토
        check_jobeffect = ('없' in under_jobeffect) or ('무' in under_jobeffect) or ('無' in under_jobeffect)

        # 출입국소 불일치 여부 검토
        correct_count = len(set(cleaned_jobname_above) & set(cleaned_jobname_under) & set(cleaned_jobname_jobenterance))

        # 검토 완료 승인처리 영역
        if correct_count >= 1 and check_jobeffect:
            # 출입국소 검토시 양호한 경우
            print(' - 결과 : 승인')
            driver.find_element_by_xpath('/html/body/div[1]/div[2]/button[1]').click()  # 승인 클릭
            driver.find_element_by_xpath('/html/body/div[2]/div[3]/button[1]').click()  # 승인하시겠습니까? Ok 버튼 클릭
            time.sleep(2)
            driver.find_element_by_xpath('//*[@id="alertDlg"]/div[3]/button').click()  # 승인 처리되었습니다. Close 버튼 클릭
        elif check_jobeffect is False:
            # 서비스영향 작업의 경우
            print(' - 결과 : 미승인')
            print(' - 사유 : 서비스 영향 작업')
            driver.find_element_by_xpath('/html/body/div[1]/div[2]/button[4]').click()  # 닫기 클릭
            row += 1
        elif correct_count == 0:
            # 출입국소 검토시 불량인 경우
            print(' - 결과 : 미승인')
            print(' - 사유 : 작업명과 출입국소간 국사명 불일치')
            driver.find_element_by_xpath('/html/body/div[1]/div[2]/button[4]').click()  # 닫기 클릭
            row += 1
        print('============================================================================\r\n')
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(2)
    except:
        print('!알림 : 외부망 전체 승인 완료')
        print('!알림 : 10초 후 프로그램이 종료됩니다.')
        driver.close()
        time.sleep(10)
        break

##
