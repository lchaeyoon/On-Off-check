import streamlit as st
import pandas as pd

def main():
    st.title("블로그 제목 추출 시스템")
    
    # 파일 업로드
    uploaded_file = st.file_uploader("블로그 데이터 파일을 선택하세요", type=['xlsx', 'csv'])
    
    if uploaded_file:
        # 파일 읽기
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
        
        if st.button("제목 추출"):
            # 제목 추출 처리
            # (여기에 실제 제목 추출 로직 추가)
            
            # 결과 다운로드
            st.download_button(
                label="추출 결과 다운로드",
                data=df.to_excel(index=False),
                file_name="추출결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ) 