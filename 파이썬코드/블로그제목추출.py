import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
import feedparser
import time

# 구글 스프레드시트 정보
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1ufCsVjPm1YJ6FvipTcKDuvGddVWQqJeF_6sahTpO7nk/edit#gid=1320512368"
spreadsheet_id = re.search(r"/d/(\S+)/edit", spreadsheet_url).group(1)
original_sheet_name = '업체관리'
cell_range = 'C6:F'

# Google Sheets API 인증 설정
credentials = service_account.Credentials.from_service_account_file('D:/이채윤 파일/코딩/colab-408723-89110ae33a5b.json')
scoped_credentials = credentials.with_scopes(['https://www.googleapis.com/auth/spreadsheets'])
service = build('sheets', 'v4', credentials=scoped_credentials)

# 가져올 URL 리스트 가져오기
url_range = f'{original_sheet_name}!L6:L'
url_values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=url_range).execute().get('values', [])

# 가져올 업체명 리스트 가져오기
company_names_range = f'{original_sheet_name}!B6:B'
company_names_values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=company_names_range).execute().get('values', [])

# 데이터가 있는지 확인 후 처리
if url_values and company_names_values:
    for index, (url_data, company_name_data) in enumerate(zip(url_values, company_names_values), 1):
        # 리스트가 비어 있으면 처리하지 않음
        if not url_data or not company_name_data:
            continue

        url = url_data[0]
        company_name = company_name_data[0]

        # Check if the sheet already exists
        sheet_names = [sheet['properties']['title'] for sheet in service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()['sheets']]
        new_sheet_name = company_name

        if new_sheet_name not in sheet_names:
            # 새로운 시트 생성
            new_sheet_request = {
                'requests': [
                    {
                        'duplicateSheet': {
                            'sourceSheetId': service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()['sheets'][0]['properties']['sheetId'],
                            'insertSheetIndex': 1,
                            'newSheetName': new_sheet_name
                        }
                    }
                ]
            }
            service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=new_sheet_request).execute()
            print(f"새로운 시트 '{new_sheet_name}'가 생성되었습니다.")

        # RSS 피드 가져오기
        feed = feedparser.parse(url)

        # RSS 피드 아이템에서 필요한 정보 추출
        max_items_to_fetch = 50
        items = feed.get("items", [])[:max_items_to_fetch]

        data = []
        for item in items:
            title = item.get("title", "")
            author = item.get("author", "")
            link = item.get("link", "")
            published = item.get("published", "")
            
            # 필요한 정보가 없을 경우 빈 문자열로 처리
            data.append([title, author, link, published])

        # Google Sheets에 값 입력
        body = {"values": data}
        service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=f'{new_sheet_name}!{cell_range}', body=body, valueInputOption='USER_ENTERED').execute()
        print(f"새로운 시트 '{new_sheet_name}'에 데이터가 입력되었습니다.")

else:
    print("No data found.")