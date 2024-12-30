from google.oauth2 import service_account
import gspread
import re
import time
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import requests
from pathlib import Path

# 블로그 순위 가져오는 함수
def get_blog_rank(driver, keyword, target_blog_id):
    if keyword == "":
        return "Invalid keyword"  # 잘못된 입력

    search_url = f"https://search.naver.com/search.naver?ssc=tab.blog.all&sm=tab_jum&query={keyword}"
    driver.get(search_url)

    time.sleep(2)  # 페이지 로딩 대기

    # 블로그 링크 찾기
    links = driver.find_elements(By.XPATH, '//*[@id="main_pack"]/section/div[1]/ul/li/div/div[1]/div[2]/a[@target="_blank"]')

    # 블로그 ID를 사용하여 순위 찾기
    for i, link in enumerate(links, start=1):
        href = link.get_attribute('href')
        blog_id = extract_blog_url(href)
        if blog_id == target_blog_id:
            return str(i) if i <= 30 else "30↓"  # 순위 반환, 30 이상은 "30↓" 

    return "30↓"  # 찾을 수 없으면 "30↓"

# 블로그 URL에서 ID 추출
def extract_blog_url(naver_url):
    match = re.search(r'https://blog\.naver\.com/(.*)', naver_url)
    return match.group(1) if match else ""

def main():
    try:
        # 인증
        credentials_path = Path('D:/이채윤 파일/코딩/colab-408723-89110ae33a5b.json')
        if not credentials_path.exists():
            raise FileNotFoundError("인증 파일을 찾을 수 없습니다.")

        credentials = service_account.Credentials.from_service_account_file(str(credentials_path))
        scoped_credentials = credentials.with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
        gc = gspread.Client(auth=scoped_credentials)

        # 스프레드시트 ID 추출
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/1xeQbM2HZn6wOeMsVtkYHflPRKp2IHziqBwHQ46_F75Y/edit#gid=1506567720"
        spreadsheet_id = re.search(r"/d/(\S+)/edit", spreadsheet_url).group(1)

        # '조회시트' 선택
        ref_sheet = gc.open_by_key(spreadsheet_id).worksheet('조회시트')

        # L8:L 셀에서 시트 제목 가져오기
        sheet_titles = ref_sheet.col_values(14)[6:]  # N열에서 시트 제목 가져오기

        # Selenium WebDriver 초기화
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

        try:
            for sheet_title in sheet_titles:
                print(f"시트 '{sheet_title}' 처리 중...")
                
                # 시트 선택
                worksheet = gc.open_by_key(spreadsheet_id).worksheet(sheet_title)

                # G23 열에 이미 데이터가 있으면 스킵
                g23_value = worksheet.acell('G23').value
                if g23_value is not None and g23_value.strip() != "":
                    print(f"시트 '{sheet_title}'에 이미 데이터가 있습니다. 넘어갑니다.")
                    continue

                # A2 셀에서 타겟 블로그 ID 가져오기
                target_blog_id = worksheet.acell('A2').value
                if target_blog_id:
                    target_blog_id = target_blog_id.strip()
                else:
                    print(f"시트 '{sheet_title}'의 블로그 ID가 없습니다. 넘어갑니다.")
                    continue

                # E23:E 열에서 키워드 가져오기
                keywords = worksheet.col_values(5)[22:]

                # G23부터 결과 값을 입력하기
                for i, keyword in enumerate(keywords, start=23):
                    keyword = keyword.strip()  # 공백 제거
                    if not keyword:  # 키워드가 비어있으면 스킵
                        continue
                        
                    blog_rank = get_blog_rank(driver, keyword, target_blog_id)

                    # G 열에 순위 업데이트
                    try:
                        worksheet.update(f"G{i}", [[blog_rank]])
                        print(f"시트 '{sheet_title}'의 {i}행에 데이터 입력: {blog_rank}")
                        time.sleep(1)  # API 호출 제한 방지
                    except requests.exceptions.JSONDecodeError:
                        print(f"JSON 응답 오류 발생: 시트 '{sheet_title}'의 {i}행 데이터 업데이트 실패.")
                    except gspread.exceptions.APIError as e:
                        print(f"Google Sheets API 오류: {e}")
                        time.sleep(60)  # API 오류 시 1분 대기
                    except Exception as e:
                        print(f"예기치 못한 오류 발생: {e}")

            print("모든 시트의 순위체크가 완료되었습니다. 블로그제목추출을 시작합니다...")
            
            # 블로그제목추출.py 실행
            blog_title_script = Path("D:/이채윤 파일/코딩/파이썬코드/블로그제목추출.py")
            if blog_title_script.exists():
                subprocess.run(["python", str(blog_title_script)])
            else:
                print("블로그제목추출.py 파일을 찾을 수 없습니다.")

        finally:
            driver.quit()  # WebDriver 종료

    except Exception as e:
        print(f"프로그램 실행 중 오류 발생: {e}")

if __name__ == "__main__":
    main()