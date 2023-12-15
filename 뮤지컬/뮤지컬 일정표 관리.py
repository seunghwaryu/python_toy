from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

# 일정표 파일이 저장될 위치
file_path = 'D:/Me/뮤지컬 스케줄/%s.xlsx'

# 인터파크에서 공연 상품페이지로 들어가는 함수
def visitInterpark(name):
    driver = webdriver.Chrome()
    driver.get('https://tickets.interpark.com/') # 인터파크 티켓 접속
    driver.implicitly_wait(time_to_wait=5) # 로드될 때까지 대기
    
    # 검색창 찾기
    ele = driver.find_element(By.CSS_SELECTOR,'#__next > div > header > div > div._navi_p92f5_16 > div > div._biSearch_p92f5_76 > div._wrap_1iig7_1 > div > input[type=text]') 
    ele.send_keys(name) # 검색어 입력
    # 검색 버튼 클릭
    driver.find_element(By.CSS_SELECTOR,'#__next > div > header > div > div._navi_p92f5_16._autoComplete_p92f5_38 > div > div._biSearch_p92f5_76 > div._wrap_1iig7_1 > div._searchInput_1iig7_16._active_1iig7_33 > button._searchBtn_1iig7_101').click() 
    # 검색해서 나오는 첫번째 공연 클릭
    driver.find_element(By.CSS_SELECTOR,'#__next > div > main > div > div > div.result-ticket_wrapper__H41_U > div.result-ticket_listWrapper__xcEo3 > a:nth-child(1)').click()
    # driver.find_element(By.XPATH,'//*[@id="__next"]/div/main/div/div/div[1]/div[2]/a[1]').click()

    # 현재 창의 핸들을 저장
    main_window = driver.current_window_handle
    driver.close() # 검색 페이지 닫기

    # 검색 페이지 창에서 공연 상품페이지 창으로 전환
    for handle in driver.window_handles:
        if handle != main_window:
            driver.switch_to.window(handle)
            break
    
    # 여기서부터 공연 상품페이지에서 작업 이루어짐
    driver.implicitly_wait(time_to_wait=5) # 로드될 때까지 대기
    try:
        driver.find_element(By.CSS_SELECTOR,'#popup-prdGuide > div > div.popupFooter > button').click() # 팝업 창 있으면 닫기
    except:
        pass
    
    return driver
        
    
# 인터파크에서 전체 일정표 가져오는 함수
def getScheduleFromWeb(name):
    driver = visitInterpark(name)
    driver.implicitly_wait(time_to_wait=5) # 로드될 때까지 대기
    driver.find_element(By.CSS_SELECTOR,'#productMainBody > nav > div > div > ul > li:nth-child(2) > a').click() # 캐스트정보 클릭
    driver.implicitly_wait(time_to_wait=5) # 로드될 때까지 대기
    tbody = driver.find_element(By.CSS_SELECTOR,'#productMainBody > div > div > div.castingDetailResult > table > tbody') # 일정표 전체 얻기

    rows = tbody.find_elements(By.TAG_NAME,'tr') # 표의 열 얻기
    columns = [r.text for r in rows[0].find_elements(By.TAG_NAME,'th')] # 표의 열 이름 저장
    schedule_info = []

    # 일정표의 정보 가져와서 2차원 배열로 저장
    for r in rows[1::]: 
        values = [e.text for e in r.find_elements(By.TAG_NAME,'td')]
        schedule_info.append(values)
    driver.close()

    df = pd.DataFrame(data=schedule_info,columns=columns) # 엑셀로 저장할 수 있게 dataframe으로 만들어주기
    
    return df

# 시트 디자인 설정 함수
def setExcelStyle(ws, need_width):
    default_width = ws.sheet_format.defaultColWidth # 기본 열의 폭 구하기
    if not default_width:
        default_width = 8
        
    cell_color = 'EBF1DE' # 셀 색깔 
    column_color = '9BBB59' # column 색깔
    set_cell_fill = PatternFill(start_color=cell_color, end_color=cell_color,fill_type='solid')
    set_column_fill = PatternFill(start_color=column_color, end_color=column_color,fill_type='solid')
    box = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin')) # 테두리 지정
    
    if(need_width > default_width):
        for i in range(2,ws.max_column):
            c = chr(65+i)
            ws.column_dimensions[c].width = need_width
    
    # 모든 셀에 가운데 정렬 및 테두리 설정
    for r in range(ws.max_row):
        for c in range(ws.max_column):
            cell = ws.cell(row=r+1,column=c+1)
            cell.alignment = Alignment(horizontal='center')
            cell.border = box
    
    # column 행에 색 채우기 및 볼드체 설정
    for c in range(ws.max_column):
        cell = ws.cell(row=1,column=c+1)
        cell.font = Font(bold=True)
        cell.fill = set_column_fill
    
    # 짝수번 행에 색 채우기    
    for r in range(1, ws.max_row, 2):
        for c in range(ws.max_column):
            cell = ws.cell(row=r+1,column=c+1)
            cell.fill = set_cell_fill
    
    return ws

# 엑셀파일 불러오는 함수
def loadExcelfile(file_name):
    # file_name과 동일한 이름의 파일이 있으면 그 파일을 불러와 데이터프레임으로 변환한 것을 반환
    # 파일이 없으면 빈 데이터프레임을 반환한다.
    # 메인함수에서의 활용을 위해 pandas로 불러온다.
    try:
        df = pd.read_excel(file_path %file_name)
    except:
        df = pd.DataFrame()
    return df

# 일정표를 엑셀파일로 저장하는 함수
def saveScheduleInExcel(schedule_df, file_name, sheet_name):
    # sheet_name이 전체스케줄이라는 것은 현재 파일이 존재하지 않는 것이다.
    if sheet_name == '전체스케줄':
        wb = Workbook() 
        wb.remove(wb['Sheet'])
    else:
        wb = load_workbook(file_path %file_name)
        
    wb.create_sheet(sheet_name) # 새로운 시트 생성
    ws = wb[sheet_name] # 앞으로 다루는 시트를 앞서 생성한 시트로 설정
    
    need_width = max([len(c) for c in schedule_df.columns])*2 # columns의 최대글자수에 폭을 맞추기 위한 너비 
    
    # 시트에 일정표 데이터 추가    
    for r in dataframe_to_rows(schedule_df, index=False, header=True):
        ws.append(r)
        
    ws = setExcelStyle(ws, need_width) # 시트 스타일 설정
    wb.save(file_path %file_name) # 엑셀파일로 저장하기

# 원하는 배역:배우 입력받는 함수
def inputCast(cast_list):
    cast_dict = dict()
    print('\n원하는 배우와 그 배우의 배역명을 입력해주세요. 입력 종료를 원하면 0을 입력해주세요. ex) 장발장 민우혁')
    print("배역명 목록: ",end='')
    print(*cast_list,sep=', ')
    print('-' * 100)
    
    while True:
        cast = input('입력: ').split()
        if(cast[0] == '0'):
            break
        cast_name = ' '.join(cast[:-1:])
        cast_dict[cast_name] = cast[-1]
    
    return cast_dict

# 일정표 필터링 하는 함수
def fillterSchdeuleByActor(cast_dict, file_name):
    df = pd.read_excel(file_path %file_name) # 파일 불러오기
    # 필터링
    for name, actor in cast_dict.items(): 
        df = df[df[name] == actor] # 필터링한 데이터프레임을 그대로 다시 저장
    
    return df

# 날짜 비교 함수
def compareDate(str_date):
    today = datetime.now()
    last_date = datetime.strptime(str_date,'%m/%d') # 문자열을 datetime 형식으로 바꿔주기
    
    # last_date는 년도가 설정되있지 않기에 현재의 월과 last_date의 월을 비교하기 위해 월의 값을 추출
    # last_date의 년도 설정을 위해 현재 년도 추출
    now_year = int(today.strftime('%Y'))
    now_month = int(today.strftime('%m'))
    date_month = int(last_date.strftime('%m'))
    
    # 뮤지컬은 보통 3개월 정도 하기 때문에 현재 10월이고 1월인 일정이 있다면 그 1월은 내년 1월을 뜻하기에 다음과 같이 년도 설정
    if now_month > 10 and date_month < 3:
        last_date = last_date.replace(year = now_year + 1)
    else:
        last_date = last_date.replace(year = now_year)
    
    # 메인에서 활용하기 위해 True와 False를 반환
    if today > last_date:
        return True
    else:
        return False
    
    
# 인터파크에서 가격 정보 가져오는 함수
def getPriceFromInterpark(name):
    driver = visitInterpark(name)
    # 전체가격버튼 클릭
    driver.find_element(By.CSS_SELECTOR,'#container > div.contents > div.productWrapper > div.productMain > div.productMainTop > div > div.summaryBody > ul > li.infoItem.infoPrice > div > ul > li.infoPriceItem.is-largePrice > a').click()
    # 가격정보 전체 얻기
    tbody = driver.find_element(By.CSS_SELECTOR,'#popup-info-price > div > div.popupBody > div > div > table > tbody')
    rows = tbody.find_elements(By.TAG_NAME,'tr') # 표의 열 얻기
    price_list = []
    for r in rows: 
            values = [e.text for e in r.find_elements(By.TAG_NAME,'td')]
            price_list.append(values)
    driver.close()
    df = pd.DataFrame(data=price_list)
    
    seat = ""
    price_info = dict()
    for i in range(len(df)):
        n = df[0][i]
        p = df[1][i]
        op = df[2][i]
        if op: # 정가일 경우
            seat = n
            if p in price_info:
                price_info[p].append([seat,op])
            else:
                price_info[p] = [[seat,op]]
        else: # 할인가 일 경우
            if n in price_info:
                price_info[n].append([seat,p])
            else:
                price_info[n] = [[seat,p]]  
                
    discount_list = list(price_info.keys())
    seat_and_price = list(price_info.values())
    price_info = {'할인명':discount_list}
    
    # 할인명에 따라 좌석등급별 금액으로 정리하기
    for layer1 in seat_and_price:
        for inp in layer1:
            s = inp[0]
            p = inp[1]
            if s in price_info:
                price_info[s].append(p)
            else:
                price_info[s] = [p]
    df = pd.DataFrame(price_info)
    df['할인명'] = df['할인명'].apply(lambda x: x.split('\n')[0]) # 할인명에 있는 날짜 제거
    df = df.drop_duplicates(ignore_index = True) # 할인명 중복 제거
    
    return df
    
# 메인 함수
input_name = input('정보를 검색할 공연명을 입력해주세요: ')
choice = int(input('1.할인정보 2.일정표: '))

if(choice == 1):
    pass
elif(choice == 2):
    schedule_df = loadExcelfile(input_name)
    
    # 입력받은 공연명과 이름이 동이한 엑셀파일이 존재 하지않다면
    if schedule_df.empty:
        schedule_df = getScheduleFromWeb(input_name)
        saveScheduleInExcel(schedule_df,input_name,'전체스케줄')
    else:
        last_date = schedule_df.iloc[-1][0] # 일정표의 마지막 일정의 날짜 가져오기
        last_date = last_date.split('(')[0] # 요일 정보 삭제
        
        # 현재 가지고 있는 일정표의 마지막 날짜가 현재보다 과거라면
        if compareDate(last_date):
            schedule_df = getScheduleFromWeb(input_name)
            saveScheduleInExcel(schedule_df,input_name,'전체스케줄')
    
    # 캐스트 입력받기
    cast_list = list(schedule_df.columns[2:])
    cast_dict = inputCast(cast_list)
    cast_dict = dict(sorted(cast_dict.items(),key= lambda item:item[1])) # 배우 이름 순으로 정렬
    
    # 필터링하기
    filtered_schedule_df = fillterSchdeuleByActor(cast_dict, input_name)
    
    # 필터링한 일정표 엑셀로 저장
    sheet_name = ','.join(list(cast_dict.values())) # 시트 이름 배우들 이름으로 설정
    saveScheduleInExcel(filtered_schedule_df,input_name,sheet_name)