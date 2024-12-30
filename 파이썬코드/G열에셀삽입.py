import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime
import locale
import gspread
import re
import subprocess
import time

# Set the locale to Korean
locale.setlocale(locale.LC_TIME, 'ko_KR')

# 인증 키 파일의 경로
credentials = service_account.Credentials.from_service_account_file('D:/이채윤 파일/코딩/colab-408723-89110ae33a5b.json')
scoped_credentials = credentials.with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
gc = gspread.Client(auth=scoped_credentials)

# Google Sheets 링크에서 시트의 고유 ID(gid) 추출
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1xeQbM2HZn6wOeMsVtkYHflPRKp2IHziqBwHQ46_F75Y/edit#gid=1506567720"
spreadsheet_id = re.search(r"/d/(\S+)/edit", spreadsheet_url).group(1)

# '조회시트' 시트 선택
ref_sheet = gc.open_by_key(spreadsheet_id).worksheet('조회시트')

# B2:B 셀에서 실행할 시트 제목 가져오기
sheet_titles = ref_sheet.col_values(12)[6:]

# Build the service
service = build('sheets', 'v4', credentials=scoped_credentials)

# 오늘 날짜 가져오기 ('YYYY-MM-DD 요일' 형식)
today_date = datetime.today().strftime('%Y-%m-%d %a')

# G1 셀에 이미 오늘 날짜가 입력된 시트는 건너뛰도록 하기 위해 G1 셀 값 가져오기
ranges = [f"'{sheet_title}'!G1" for sheet_title in sheet_titles]
result = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges).execute()

# 모든 시트에 대한 작업 완료 여부 플래그
all_sheets_completed = True

# 시트 G1의 값을 미리 확인한 후 처리
for i, sheet_title in enumerate(sheet_titles):
    try:
        # G1 셀에 오늘 날짜가 이미 있는 경우 건너뛰기
        g1_value = result.get('valueRanges', [])[i].get('values', [[None]])[0][0]
        if g1_value == today_date:
            print(f"시트 '{sheet_title}'는 이미 오늘 날짜로 업데이트되어 있습니다. 다음 시트로 넘어갑니다.")
            continue
        
        # 시트 선택
        worksheet = gc.open_by_key(spreadsheet_id).worksheet(sheet_title)
        
        # 시트 ID 가져오기
        sheet_id = worksheet.id

        # 열을 삽입할 요청 생성 (G열, 7번째 열에)
        requests = [{
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 6,  # G열이 7번째 열이므로
                    "endIndex": 7
                }
            }
        }]

        # 열 삽입 및 G1, G2 셀 업데이트 요청을 배치로 처리
        batch_requests = {
            "requests": requests
        }

        # G1셀에 오늘 날짜 입력
        g1_update = {
            'range': f'{sheet_title}!G1',
            'values': [[today_date]]
        }
        
        # G2셀에 수식 입력
        g2_formula = 'COUNTIF(G3:G,1)+COUNTIF(G3:G,2)+COUNTIF(G3:G,3)+COUNTIF(G3:G,4)+COUNTIF(G3:G,5)+COUNTIF(G3:G,6)+COUNTIF(G3:G,7)+COUNTIF(G3:G,8)+COUNTIF(G3:G,9)+COUNTIF(G3:G,10)'
        g2_update = {
            'range': f'{sheet_title}!G2',
            'values': [[f'={g2_formula}']]
        }
        
        # 값 업데이트 요청 처리 (G1, G2)
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_requests).execute()
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body={
            'valueInputOption': 'USER_ENTERED',
            'data': [g1_update, g2_update]
        }).execute()

        print(f"시트 '{sheet_title}'에 대한 작업이 완료되었습니다.")

    except gspread.exceptions.WorksheetNotFound:
        print(f"시트 '{sheet_title}'을(를) 찾을 수 없습니다. 건너뜁니다.")
        all_sheets_completed = False
    
    except gspread.exceptions.APIError as e:
        print(f"API Error for '{sheet_title}': {e}. 10초 후 다시 시도합니다.")
        time.sleep(10)  # 잠시 대기한 후 다시 시도
        all_sheets_completed = False

# 모든 시트에 대한 작업이 완료되었다면 B 실행 파일 시작
if all_sheets_completed:
    subprocess.run(["python", r"D:/이채윤 파일/코딩/파이썬코드/순위체크.py"])