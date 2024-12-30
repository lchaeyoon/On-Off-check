import streamlit as st
import pandas as pd

def main():
    st.title("엑셀 G열 삽입 시스템")
    
    # 파일 업로드
    uploaded_file = st.file_uploader("엑셀 파일을 선택하세요", type=['xlsx', 'xls'])
    
    if uploaded_file:
        # 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file)
        
        # G열에 삽입할 내용 입력
        insert_text = st.text_input("G열에 삽입할 내용을 입력하세요")
        
        if st.button("G열 삽입"):
            # G열에 내용 삽입
            df.insert(6, 'G열', insert_text)
            
            # 결과 파일 다운로드
            st.download_button(
                label="결과 파일 다운로드",
                data=df.to_excel(index=False),
                file_name="결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ) 