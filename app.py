import streamlit as st
import pandas as pd

st.set_page_config(page_title="테스트 앱", layout="wide")

st.title("🔋 Streamlit 업로드 테스트")
st.write("이 문장이 보이면 앱은 정상 실행 중이야.")

uploaded_file = st.file_uploader(
    "여기에 2주차 (6).xlsx 파일을 업로드해줘",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("엑셀 파일을 업로드해줘.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file)
    st.success("파일 업로드 성공!")
    st.write("시트 목록:")
    st.write(xls.sheet_names)

    selected_sheet = st.selectbox("시트 선택", xls.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
    st.dataframe(df.head(20), use_container_width=True)

except Exception as e:
    st.error(f"파일 읽기 오류: {e}")

