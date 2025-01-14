import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import requests
import json
import socket
import platform
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Windows 환경에서만 import
if os.name == 'nt':
    import win32evtlog
    import win32evtlogutil
    import win32con

def load_custom_holidays():
    """자체 휴가 목록 로드"""
    custom_holiday_file = 'custom_holidays.csv'
    try:
        if os.path.exists(custom_holiday_file):
            df = pd.read_csv(custom_holiday_file)
            return dict(zip(df['date'], df['description']))
        return {}
    except Exception:
        return {}

def save_custom_holiday(date, description):
    """자체 휴가 추가"""
    custom_holiday_file = 'custom_holidays.csv'
    try:
        if os.path.exists(custom_holiday_file):
            df = pd.read_csv(custom_holiday_file)
        else:
            df = pd.DataFrame(columns=['date', 'description'])
        
        # 이미 존재하는 날짜인지 확인
        if date not in df['date'].values:
            new_row = pd.DataFrame({'date': [date], 'description': [description]})
            df = pd.concat([df, new_row], ignore_index=True)
            df.to_csv(custom_holiday_file, index=False)
            return True
        return False
    except Exception:
        return False

def delete_custom_holiday(date):
    """자체 휴가 삭제"""
    custom_holiday_file = 'custom_holidays.csv'
    try:
        if os.path.exists(custom_holiday_file):
            df = pd.read_csv(custom_holiday_file)
            df = df[df['date'] != date]
            df.to_csv(custom_holiday_file, index=False)
            return True
        return False
    except Exception:
        return False

def get_holidays():
    """공휴일과 자체 휴가 목록 통합"""
    holidays = {
        # 기존 공휴일 목록
        "2024-01-01": "신정",
        "2024-02-09": "설날",
        "2024-02-10": "설날",
        "2024-02-11": "설날",
        "2024-02-12": "대체공휴일(설날)",
        "2024-03-01": "삼일절",
        "2024-04-10": "21대 총선",
        "2024-05-05": "어린이날",
        "2024-05-06": "대체공휴일(어린이날)",
        "2024-05-15": "부처님오신날",
        "2024-06-06": "현충일",
        "2024-08-15": "광복절",
        "2024-09-16": "추석",
        "2024-09-17": "추석",
        "2024-09-18": "추석",
        "2024-10-03": "개천절",
        "2024-10-09": "한글날",
        "2024-12-25": "크리스마스",
        "2024-12-26": "겨울방학",
        "2024-12-27": "겨울방학",
        "2024-12-30": "겨울방학",
        "2024-12-31": "겨울방학",    

        # 2025년 공휴일
        "2025-01-01": "신정",
        "2025-01-27": "임시공휴일",
        "2025-01-28": "설날",
        "2025-01-29": "설날",
        "2025-01-30": "설날",
        "2025-03-01": "삼일절",
        "2025-05-05": "어린이날",
        "2025-05-06": "부처님오신날",
        "2025-06-06": "현충일",
        "2025-08-15": "광복절",
        "2025-10-03": "개천절",
        "2025-10-09": "한글날",
        "2025-12-25": "크리스마스"
    }
    
    # 자체 휴가 추가
    custom_holidays = load_custom_holidays()
    holidays.update(custom_holidays)
    
    return holidays

def is_holiday(date):
    """공휴일 여부 확인"""
    holidays = get_holidays()
    date_str = date.strftime('%Y-%m-%d')
    return date_str in holidays

def get_holiday_name(date):
    """공휴일 이름 반환"""
    holidays = get_holidays()
    date_str = date.strftime('%Y-%m-%d')
    return holidays.get(date_str)

def get_local_pc_events(start_date=None, end_date=None):
    """PC 사용 기록 반환"""
    events = []
    try:
        computer_name = platform.node()
        
        # Windows 환경인 경우 실제 이벤트 로그 사용 및 Google Sheets에 저장
        if os.name == 'nt':
            try:
                hand = win32evtlog.OpenEventLog(None, "System")
                flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
                
                while True:
                    events_raw = win32evtlog.ReadEventLog(hand, flags, 0)
                    if not events_raw:
                        break
                        
                    for event in events_raw:
                        try:
                            event_id = event.EventID & 0xFFFF
                            if event_id in [6005, 6006, 6008, 6009, 1074]:
                                event_date = event.TimeGenerated.replace(tzinfo=None)
                                
                                if start_date and end_date:
                                    start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                                    end_dt = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
                                    if not (start_dt <= event_date <= end_dt):
                                        continue
                                
                                event_type = '시작' if event_id in [6005, 6009] else '종료'
                                events.append({
                                    'time': event_date,
                                    'type': event_type,
                                    'event_id': event_id,
                                    'computer': computer_name
                                })
                        except Exception:
                            continue
                            
                win32evtlog.CloseEventLog(hand)
                
                # 이벤트를 Google Sheets에 저장
                if events:
                    save_events_to_sheet(events)
                
            except Exception as e:
                st.error(f"이벤트 로그 접근 오류: {str(e)}")
                return []
        
        # 모든 환경에서 Google Sheets에서 데이터 읽기
        events = load_events_from_sheet(start_date, end_date, computer_name)
                    
    except Exception as e:
        st.error(f"이벤트 생성 중 오류 발생: {str(e)}")
    
    return sorted(events, key=lambda x: x['time'], reverse=True)

def save_events_to_sheet(events):
    """이벤트를 Google Sheets에 저장"""
    try:
        service_account_info = json.loads(st.secrets["google_service_account"])
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )
        
        service = build('sheets', 'v4', credentials=creds)
        SPREADSHEET_ID = '1-xF7-9VK3Ty5-ARnp0RSqyzrYJXmhW1phaPZTX42SLs'
        SHEET_NAME = 'PC_Events'  # 새로운 시트
        
        values = [[
            event['time'].strftime('%Y-%m-%d %H:%M:%S'),
            event['type'],
            event['event_id'],
            event['computer']
        ] for event in events]
        
        body = {
            'values': values
        }
        
        service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{SHEET_NAME}!A:D',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        
    except Exception as e:
        st.error(f"Google Sheets 저장 오류: {str(e)}")

def load_events_from_sheet(start_date=None, end_date=None, computer_name=None):
    """Google Sheets에서 이벤트 로드"""
    try:
        service_account_info = json.loads(st.secrets["google_service_account"])
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )
        
        service = build('sheets', 'v4', credentials=creds)
        SPREADSHEET_ID = '1-xF7-9VK3Ty5-ARnp0RSqyzrYJXmhW1phaPZTX42SLs'
        SHEET_NAME = 'PC_Events'
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f'{SHEET_NAME}!A:D'
        ).execute()
        
        values = result.get('values', [])
        events = []
        
        for row in values:
            if len(row) >= 4:
                event_time = datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                if computer_name and row[3] != computer_name:
                    continue
                    
                if start_date and end_date:
                    start_dt = datetime.strptime(start_date, '%Y-%m-%d')
                    end_dt = datetime.strptime(end_date, '%Y-%m-%d') + timedelta(days=1)
                    if not (start_dt <= event_time <= end_dt):
                        continue
                
                events.append({
                    'time': event_time,
                    'type': row[1],
                    'event_id': int(row[2]),
                    'computer': row[3]
                })
        
        return events
        
    except Exception as e:
        st.error(f"Google Sheets 로드 오류: {str(e)}")
        return []

def format_hours_to_time(hours):
    """시간을 HH:MM 형식으로 변환"""
    if isinstance(hours, str):
        if hours in ['-', '']:
            return hours
        try:
            hours = float(hours.replace('시간', ''))
        except ValueError:
            return hours
    
    total_minutes = int(hours * 60)
    hours = total_minutes // 60
    minutes = total_minutes % 60
    
    # 두 자리 수로 포맷팅
    return f"{hours:02d}:{minutes:02d}"

def calculate_work_hours(start_time, end_time, date):
    """근무시간 계산 (점심시간 제외)"""
    # 공휴일 체크
    if is_holiday(date):
        holiday_name = get_holiday_name(date)
        # 자체 휴가인지 확인
        custom_holidays = load_custom_holidays()
        date_str = date.strftime('%Y-%m-%d')
        if date_str in custom_holidays:
            return "", custom_holidays[date_str]  # 자체 휴가는 근무시간 공란, 설명 표시
        return "8:00", f"{holiday_name}"  # 공휴일은 8시간으로 표시
    
    # 주말 체크
    if date.weekday() >= 5:  # 5: 토요일, 6: 일요일
        return "-", "주말"  # 주말은 근무시간 표시하지 않음
        
    if not start_time or not end_time:
        return "-", ""
        
    total_seconds = (end_time - start_time).total_seconds()
    
    # 점심시간 (12:00-13:00) 제외
    lunch_start = start_time.replace(hour=12, minute=0, second=0)
    lunch_end = start_time.replace(hour=13, minute=0, second=0)
    
    # 근무시간이 점심심심시간을 포포함하는 경우
    if start_time <= lunch_start and end_time >= lunch_end:
        total_seconds -= 3600  # 1시간(3600초) 제외
    
    hours = max(0, total_seconds / 3600)  # 음수 방지
    return format_hours_to_time(hours), ""

def get_date_range(start_date, end_date):
    """시작일부터 종료일까지의 평일 목록 반환 (토/일 제외)"""
    date_list = []
    current = datetime.strptime(start_date, '%Y-%m-%d')
    end = datetime.strptime(end_date, '%Y-%m-%d')
    
    while current <= end:
        # 평일만 추가 (weekday: 0=월요일, 5=토요일, 6=일요일)
        if current.weekday() < 5:
            date_list.append(current.strftime('%Y-%m-%d'))
        current += timedelta(days=1)
    
    return date_list

def get_week_range(date):
    """월~금 기준 주차 시작일과 종료일 반환"""
    date_obj = datetime.strptime(date, '%Y-%m-%d')
    
    # 해당 날짜의 요일 (0: 월요일, 6: 일요일)
    weekday = date_obj.weekday()
    
    # 월요일부터 금요일까지의 날짜 계산
    monday = date_obj - timedelta(days=weekday)  # 해당 주의 월요일
    friday = monday + timedelta(days=4)  # 해당 주의 금요일
    
    return monday, friday

def calculate_weekly_stats(daily_records):
    """주차별 통계 계산"""
    weekly_stats = {}
    
    for record in daily_records.values():
        date = record['날짜']
        monday, friday = get_week_range(date)
        
        # 연도와 월이 바뀌는 경우 처리
        if monday.month != friday.month:
            if monday.year != friday.year:
                week_key = f"{friday.strftime('%Y-%m')} 1주차"
            else:
                first_day_of_month = friday.replace(day=1)
                first_monday = first_day_of_month - timedelta(days=first_day_of_month.weekday())
                week_num = ((friday - first_monday).days // 7) + 1
                week_key = f"{friday.strftime('%Y-%m')} {week_num}주차"
        else:
            first_day_of_month = monday.replace(day=1)
            first_monday = first_day_of_month - timedelta(days=first_day_of_month.weekday())
            week_num = ((monday - first_monday).days // 7) + 1
            week_key = f"{monday.strftime('%Y-%m')} {week_num}주차"
        
        if week_key not in weekly_stats:
            weekly_stats[week_key] = {
                '근무시간': [],
                '시작': monday.strftime('%Y-%m-%d'),
                '종료': friday.strftime('%Y-%m-%d')
            }
        
        # 근무시간이 있고 숫자로 변환 가능한 경우만 추가
        if record['근무시간'] and record['근무시간'] not in ['-', '']:
            try:
                # HH:MM 형식을 시간으로 변환
                hours, minutes = map(int, record['근무시간'].split(':'))
                hours = hours + minutes / 60
                weekly_stats[week_key]['근무시간'].append(hours)
            except ValueError:
                continue
    
    # 통계 계산
    result = []
    for week, data in weekly_stats.items():
        if data['근무시간']:
            total_hours = sum(data['근무시간'])
            avg_hours = total_hours / len(data['근무시간'])
            result.append({
                '주차': week,
                '기간': f"{data['시작']} ~ {data['종료']}",
                '평균 근무시간': format_hours_to_time(avg_hours),
                '총 근무시간': format_hours_to_time(total_hours),
                '근무일수': len(data['근무시간'])
            })
    
    return sorted(result, key=lambda x: x['주차'], reverse=True)

def get_computer_info():
    """PC 정보 반환"""
    try:
        return platform.node()  # 실제 PC 장치명 반환
    except:
        return "알 수 없음"

def update_google_sheet(records, employee_name):
    """구글 시트 업데이트 함수 수정"""
    try:
        # 서비스 계정 키를 환경 변수에서 가져오기
        service_account_info = json.loads(st.secrets["google_service_account"])
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES
        )
        
        # 구글 시트 API 클라이언트 생성
        service = build('sheets', 'v4', credentials=creds)
        
        # 스프레드시트 ID와 범위 지정
        SPREADSHEET_ID = '1-xF7-9VK3Ty5-ARnp0RSqyzrYJXmhW1phaPZTX42SLs'
        SHEET_NAME = '출퇴근관'
        START_ROW = 5  # 데이터 시작 행
        
        # 데이터 포맷팅
        values = []
        computer_name = get_computer_info()
        
        # 각 레코드의 주차차 정보보 계산
        for record in records:
            date = record.get('날짜', '')
            monday, friday = get_week_range(date)
            
            # 연도와 월이 바뀌는 경우 처리
            if monday.month != friday.month:
                if monday.year != friday.year:
                    week_info = f"{friday.strftime('%Y-%m')} 1주차"
                else:
                    first_day_of_month = friday.replace(day=1)
                    first_monday = first_day_of_month - timedelta(days=first_day_of_month.weekday())
                    week_num = ((friday - first_monday).days // 7) + 1
                    week_info = f"{friday.strftime('%Y-%m')} {week_num}주차"
            else:
                first_day_of_month = monday.replace(day=1)
                first_monday = first_day_of_month - timedelta(days=first_day_of_month.weekday())
                week_num = ((monday - first_monday).days // 7) + 1
                week_info = f"{monday.strftime('%Y-%m')} {week_num}주차"
            
            values.append([
                employee_name,          # 직원명
                computer_name,          # PC 정보
                week_info,              # 주차 정보
                date,                   # 날짜
                record.get('PC 시작', ''),  # PC 시작
                record.get('PC 종료', ''),  # PC 종종료
                record.get('근무시간', ''),  # 근무시간
                record.get('비고', '')   # 비고
            ])
        
        # 먼저 시트 ID 가져오기
        sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheet_id = None
        for sheet in sheet_metadata.get('sheets', ''):
            if sheet.get('properties', {}).get('title') == SHEET_NAME:
                sheet_id = sheet.get('properties', {}).get('sheetId')
                break
        
        if sheet_id is None:
            return False, "지정된 시트를 찾을 수 없습니다."
        
        if values:
            # 1. 새로운 행 삽입
            insert_request = {
                'requests': [{
                    'insertRange': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': START_ROW - 1,
                            'endRowIndex': START_ROW - 1 + len(values)
                        },
                        'shiftDimension': 'ROWS'
                    }
                }]
            }
            
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=insert_request
            ).execute()
            
            # 2. 삽입된 행에 데이터 입력 (8개 컬럼으로 수정)
            range_name = f'{SHEET_NAME}!A{START_ROW}:H{START_ROW + len(values) - 1}'
            body = {
                'values': values
            }
            
            result = service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            return True, f"{len(values)}개의 데이터가 구글 시트에 입력되었습니다."
        
        return True, "입력할 데이터가 없습니다."
        
    except Exception as e:
        return False, f"구글 시트 업데이트 중 오류 발생: {str(e)}"

def format_time(time_str):
    """시간을 HH:MM 형식으로 변환"""
    if not time_str or time_str == '-':
        return time_str
    try:
        time_obj = datetime.strptime(time_str, '%H:%M')
        return time_obj.strftime('%H:%M')
    except ValueError:
        return time_str

def main():
    st.set_page_config(page_title="PC 사용 기록 시스템", page_icon="🖥️", layout="wide")
    st.title("🖥️ PC 사용 기록 시스템")
    
    # PC 정보 표시
    computer_name = get_computer_info()
    st.caption(f"📌 현재 PC: {computer_name}")
    
    # 사이드바에 직원명과 날짜 선택
    with st.sidebar:
        st.header("👤 직원 정보")
        employee_name = st.text_input("직원명", placeholder="직원명을 입력하세요")
        
        if employee_name:  # 직원명이 입력된 경우에만 날짜 선택 표시
            st.markdown("---")
            st.header("📅 날짜 선택")
            
            # 오늘 날짜와 어제 날짜 계산
            today = datetime.now().date()
            yesterday = today - timedelta(days=1)
            
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input(
                    "시작일", 
                    yesterday - timedelta(days=6),  # 기본값: 7일 전
                    format="YYYY-MM-DD",
                    max_value=yesterday  # 최대 선택 가능 날짜: 어제
                ).strftime('%Y-%m-%d')
            with col2:
                end_date = st.date_input(
                    "종료일", 
                    yesterday,  # 기본값: 어제
                    format="YYYY-MM-DD",
                    max_value=yesterday  # 최대 선택 가능 날짜: 어제
                ).strftime('%Y-%m-%d')
    
    # 직원명이 입력된 경우에만 데이터 조회 및 표시
    if employee_name:
        # 메인 화면
        events = get_local_pc_events(start_date, end_date)
        computer_name = get_computer_info()  # 현재 PC 정보 가져오기
        
        # 선택한 기간의 모든 평일 가져오기
        all_workdays = get_date_range(start_date, end_date)
        
        if all_workdays:
            daily_records = {}
            
            # 먼저 모든 평일에 대한 기본 레코드 생성
            for date in all_workdays:
                current_date = datetime.strptime(date, '%Y-%m-%d')
                
                # 공휴일 체크를 여기서 먼저 수행
                if is_holiday(current_date):
                    work_hours, note = calculate_work_hours(None, None, current_date)
                    daily_records[date] = {
                        '날짜': date,
                        'PC 시작': '',
                        'PC 종료': '',
                        '근무시간': work_hours,
                        '비고': note,
                        'PC': computer_name  # PC 정보 추가
                    }
                else:
                    daily_records[date] = {
                        '날짜': date,
                        'PC 시작': '',
                        'PC 종료': '',
                        '근무시간': '',
                        '비고': '',
                        'PC': computer_name  # PC 정보 추가
                    }
            
            # 이벤트 데이터로 레코드 업데이트
            for event in events:
                date = event['time'].strftime('%Y-%m-%d')
                if date in daily_records and not is_holiday(event['time']):  # 공휴일이 아닌 경우만 처리
                    # 평일인 경우 정상적으로 PC 시작/종료 시간 기록
                    time_str = format_time(event['time'].strftime('%H:%M'))  # 시간 형식 변환환환
                    if event['type'] == '시작':
                        daily_records[date]['PC 시작'] = time_str
                    else:
                        daily_records[date]['PC 종료'] = time_str
                    
                    # 근무시간 계산 (PC 시작/종료 시간이 모두 있을 때만)
                    if daily_records[date]['PC 시작'] and daily_records[date]['PC 종료']:  # 빈 문자열 체크
                        start_time = datetime.strptime(f"{date} {daily_records[date]['PC 시작']}", '%Y-%m-%d %H:%M')
                        end_time = datetime.strptime(f"{date} {daily_records[date]['PC 종료']}", '%Y-%m-%d %H:%M')
                        work_hours, note = calculate_work_hours(start_time, end_time, start_time)
                        daily_records[date]['근무시간'] = work_hours
                        daily_records[date]['비고'] = note
            
            # 데이터프레임 생성 (날짜 순으로 정렬)
            df = pd.DataFrame(daily_records.values())
            df = df.drop(columns=['PC'])  # PC 컬럼 제거
            df = df.sort_values('날짜', ascending=False)
            
            # 테이블 표시
            st.markdown(f"### 📊 {employee_name}님의 일자별 PC 사용 기록")
            st.dataframe(
                df,
                column_config={
                    "날짜": st.column_config.TextColumn("날짜", width=100),
                    "PC 시작": st.column_config.TextColumn("PC 시작", width=100),
                    "PC 종료": st.column_config.TextColumn("PC 종료", width=100),
                    "근무시간": st.column_config.TextColumn("근무시간", width=100),
                    "비고": st.column_config.TextColumn("비고", width=150)
                },
                hide_index=True
            )
            
            # 주차별 통계 표시
            st.markdown(f"### 📅 {employee_name}님의 주차별 근무 통계")
            weekly_stats = calculate_weekly_stats(daily_records)
            if weekly_stats:
                weekly_df = pd.DataFrame(weekly_stats)
                st.dataframe(
                    weekly_df,
                    column_config={
                        "주차": st.column_config.TextColumn("주차", width=120),
                        "기간": st.column_config.TextColumn("기간", width=200),
                        "평균 근무시간": st.column_config.TextColumn("평균 근무시간", width=120),
                        "총 근무시간": st.column_config.TextColumn("총 근무시간", width=120),
                        "근무일수": st.column_config.NumberColumn("근무일수", width=100)
                    },
                    hide_index=True
                )
            else:
                st.info("주차별 통계를 계산할 수 있는 근무 기록이 없습니다.")
            
            # 구글 시트에 데이터 입력 버튼
            if st.button("📊 구글 시트에 데이터 입력", use_container_width=True):
                success, message = update_google_sheet(df.to_dict('records'), employee_name)
                if success:
                    st.success(message)
                else:
                    st.error(message)
        else:
            st.info("👀 선택한 기간에 PC 사용 기록이 없습니다.")
    else:
        # 직원명 미입력 시 안내 메시지
        st.warning("👆 직원명을 입력해주세요.")

    # 사이드바에 휴가 관리 섹션 추가
    with st.sidebar:
        st.markdown("---")
        st.header("🏖️ 휴가 관리")
        
        # 휴가 등록
        with st.expander("휴가 등록"):
            holiday_date = st.date_input(
                "휴가 날짜",
                datetime.now(),
                format="YYYY-MM-DD"
            )
            holiday_desc = st.text_input("휴가 설명", placeholder="예: 연차, 반차, 교육 등")
            
            if st.button("등록", use_container_width=True):
                date_str = holiday_date.strftime('%Y-%m-%d')
                if save_custom_holiday(date_str, holiday_desc):
                    st.success("휴가가 등록되었습니다.")
                else:
                    st.error("이미 등록된 날짜입니다.")
        
        # 등록된 휴가 목록
        with st.expander("등록된 휴가 목록"):
            custom_holidays = load_custom_holidays()
            if custom_holidays:
                for date, desc in custom_holidays.items():
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.write(f"{date}: {desc}")
                    with col2:
                        if st.button("삭제", key=f"del_{date}"):
                            if delete_custom_holiday(date):
                                st.rerun()
            else:
                st.info("등록된 휴가가 없습니다.")

if __name__ == "__main__":
    main() 
