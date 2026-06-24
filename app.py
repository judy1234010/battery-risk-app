import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="배터리 공급망 조기경보 대시보드",
    page_icon="🔋",
    layout="wide"
)

st.title("🔋 배터리 공급망 조기경보 대시보드")
st.caption("수입구조·국가집중도·최종위험점수를 기반으로 공급망 위험을 진단합니다.")

# ---------------------------------
# 공통 함수
# ---------------------------------
def to_ym_str(series):
    return series.astype(str).str.replace(".0", "", regex=False)

def safe_num(series):
    return pd.to_numeric(series, errors="coerce")

def fmt(x, digits=2):
    if pd.isna(x):
        return "-"
    return f"{x:,.{digits}f}"

def safe_read(uploaded_file, sheet_name):
    try:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name)
    except Exception:
        return None

# ---------------------------------
# 파일 업로드
# ---------------------------------
uploaded_file = st.file_uploader(
    "여기에 2주차 (6).xlsx 파일을 업로드해줘",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("엑셀 파일을 업로드하면 대시보드가 시작돼.")
    st.stop()

# ---------------------------------
# 데이터 읽기
# ---------------------------------
@st.cache_data
def load_data(file):
    item_info = safe_read(file, "ITEM_INFO")
    country = safe_read(file, "COUNTRY_MONTHLY")
    hs_summary = safe_read(file, "HS_MONTHLY_SUMMARY")
    panel = safe_read(file, "PANEL_MONTHLY")
    alert = safe_read(file, "ALERT_RESULT")
    compare = safe_read(file, "체인별 비교표")
    entropy = safe_read(file, "ENTROPY_WEIGHT")
    return item_info, country, hs_summary, panel, alert, compare, entropy

item_info, country, hs_summary, panel, alert, compare, entropy = load_data(uploaded_file)

required = {
    "COUNTRY_MONTHLY": country,
    "HS_MONTHLY_SUMMARY": hs_summary,
    "PANEL_MONTHLY": panel,
    "ALERT_RESULT": alert,
    "체인별 비교표": compare,
    "ENTROPY_WEIGHT": entropy
}

missing = [k for k, v in required.items() if v is None]
if missing:
    st.error(f"필수 시트를 찾지 못했어: {', '.join(missing)}")
    st.stop()

# ---------------------------------
# 전처리
# ---------------------------------
for df in [country, hs_summary, panel, alert]:
    if "연월" in df.columns:
        df["연월"] = to_ym_str(df["연월"])

for df in [country, hs_summary, panel, alert]:
    if "체인구분" in df.columns:
        df["체인구분"] = df["체인구분"].astype(str)

if "국가별수입비중" in country.columns:
    country["국가별수입비중"] = safe_num(country["국가별수입비중"])

if "최종위험점수" in alert.columns:
    alert["최종위험점수"] = safe_num(alert["최종위험점수"])

if "상위1국의존도" in panel.columns:
    panel["상위1국의존도"] = safe_num(panel["상위1국의존도"])

if "상위3국집중도" in panel.columns:
    panel["상위3국집중도"] = safe_num(panel["상위3국집중도"])

if "HHI" in panel.columns:
    panel["HHI"] = safe_num(panel["HHI"])

chains = sorted(panel["체인구분"].dropna().unique().tolist())
months = sorted(panel["연월"].dropna().unique().tolist())

# ---------------------------------
# 메뉴
# ---------------------------------
menu = st.sidebar.radio(
    "메뉴 선택",
    ["조기경보 대시보드", "품목·국가 취약성 분석", "체인별 비교", "방법론 설명", "시트 미리보기"]
)

# ---------------------------------
# 1. 조기경보 대시보드
# ---------------------------------
if menu == "조기경보 대시보드":
    st.header("📊 조기경보 대시보드")

    c1, c2 = st.columns(2)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", chains)
    with c2:
        selected_month = st.selectbox("연월 선택", months, index=len(months)-1)

    panel_row = panel[
        (panel["체인구분"] == selected_chain) &
        (panel["연월"] == selected_month)
    ]

    alert_row = alert[
        (alert["체인구분"] == selected_chain) &
        (alert["연월"] == selected_month)
    ]

    if panel_row.empty or alert_row.empty:
        st.warning("선택한 조건의 데이터가 없어.")
        st.stop()

    p = panel_row.iloc[0]
    a = alert_row.iloc[0]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("최종위험점수", fmt(a.get("최종위험점수", None)))
    m2.metric("최종경보등급", str(a.get("최종경보등급", "-")))
    m3.metric("대체조달가능성", str(a.get("대체조달가능성", "-")))
    m4.metric("FTA 활용비중", fmt(a.get("fta_ratio", None)) + "%" if pd.notna(a.get("fta_ratio", None)) else "-")

    m5, m6, m7 = st.columns(3)
    m5.metric("상위1국의존도", fmt(p.get("상위1국의존도", None)) + "%" if pd.notna(p.get("상위1국의존도", None)) else "-")
    m6.metric("상위3국집중도", fmt(p.get("상위3국집중도", None)) + "%" if pd.notna(p.get("상위3국집중도", None)) else "-")
    m7.metric("HHI", fmt(p.get("HHI", None)))

    risk_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    if all(col in alert_row.columns for col in risk_cols):
        risk_df = pd.DataFrame({
            "리스크유형": ["가격", "수급", "물류", "정책이벤트"],
            "점수": [
                a["가격리스크점수"],
                a["수급리스크점수"],
                a["물류리스크점수"],
                a["정책이벤트리스크점수"]
            ]
        })

        st.markdown("### 리스크 유형별 점수")
        fig = px.bar(
            risk_df,
            x="리스크유형",
            y="점수",
            color="리스크유형",
            text_auto=".2f"
        )
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 자동 해석")
    st.write(f"- **{selected_chain}**의 **{selected_month} 기준 경보등급은 `{a.get('최종경보등급', '-')}`** 이야.")
    st.write(f"- 보정사유: **{a.get('보정사유', '-')}**")
    st.write(f"- 비고: **{a.get('비고', '-')}**")

    if pd.notna(p.get("상위1국의존도", None)):
        if p["상위1국의존도"] >= 70:
            st.write("- 상위 1개국 의존도가 매우 높아 **수입선 다변화 필요성**이 커.")
        elif p["상위1국의존도"] >= 50:
            st.write("- 상위 1개국 의존도가 높아 **집중 리스크 관리**가 필요해.")
        else:
            st.write("- 상위 1개국 의존도는 비교적 안정적인 편이야.")

    st.markdown("### 월별 최종위험점수 추이")
    chain_df = alert[alert["체인구분"] == selected_chain].copy()
    if not chain_df.empty and "최종위험점수" in chain_df.columns:
        fig2 = px.line(
            chain_df,
            x="연월",
            y="최종위험점수",
            markers=True,
            title=f"{selected_chain} 월별 최종위험점수 추이"
        )
        fig2.update_layout(height=420)
        st.plotly_chart(fig2, use_container_width=True)

# ---------------------------------
# 2. 품목·국가 취약성 분석
# ---------------------------------
elif menu == "품목·국가 취약성 분석":
    st.header("🌍 품목·국가 취약성 분석")

    c1, c2, c3 = st.columns(3)

    with c1:
        selected_chain = st.selectbox(
            "체인구분 선택",
            sorted(country["체인구분"].dropna().unique().tolist())
        )

    with c2:
        month_candidates = sorted(
            country[country["체인구분"] == selected_chain]["연월"].dropna().unique().tolist()
        )
        selected_month = st.selectbox("연월 선택", month_candidates, index=len(month_candidates)-1)

    with c3:
        hs_candidates = sorted(
            country[
                (country["체인구분"] == selected_chain) &
                (country["연월"] == selected_month)
            ]["HS코드"].astype(str).str.replace(".0", "", regex=False).dropna().unique().tolist()
        )
        selected_hs = st.selectbox("HS코드 선택", hs_candidates)

    detail_df = country[
        (country["체인구분"] == selected_chain) &
        (country["연월"] == selected_month) &
        (country["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs)
    ].copy()

    summary_df = hs_summary[
        (hs_summary["체인구분"] == selected_chain) &
        (hs_summary["연월"] == selected_month) &
        (hs_summary["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs)
    ].copy()

    if detail_df.empty:
        st.warning("선택한 조건의 데이터가 없어.")
        st.stop()

    detail_df = detail_df.sort_values("국가별수입비중", ascending=False)

    if not summary_df.empty:
        h = summary_df.iloc[0]
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("품목명", str(h.get("품목명", "-")))
        s2.metric("총수입금액", fmt(h.get("총수입금액", None), 0))
        s3.metric("상위공급국", str(h.get("상위공급국", "-")))
        s4.metric("수입국수", fmt(h.get("수입국수", None), 0))

    show_cols = [
        "국가명", "지역권", "FTA여부", "국가별수입비중", "지역권별수입비중",
        "상위공급국여부", "기본평가점수", "총보정점수", "최종보정점수", "최종판정", "비고"
    ]
    show_cols = [c for c in show_cols if c in detail_df.columns]

    st.markdown("### 국가별 취약성 상세")
    st.dataframe(detail_df[show_cols], use_container_width=True, height=420)

    if "국가명" in detail_df.columns and "국가별수입비중" in detail_df.columns:
        fig = px.bar(
            detail_df.head(10),
            x="국가명",
            y="국가별수입비중",
            color="지역권" if "지역권" in detail_df.columns else None,
            text_auto=".2f",
            title="상위 국가별 수입비중"
        )
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 해석 포인트")
    top_row = detail_df.iloc[0]
    st.write(f"- 최상위 공급국은 **{top_row.get('국가명', '-')}** 이고 수입비중은 **{fmt(top_row.get('국가별수입비중', None))}%** 야.")

    if "FTA여부" in detail_df.columns:
        fta_n = detail_df[detail_df["FTA여부"] == "N"]
        st.write(f"- FTA 미체결/미활용 국가 건수: **{len(fta_n)}건**")

    if item_info is not None and "HS코드" in item_info.columns:
        info_df = item_info[item_info["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs]
        if not info_df.empty and "선정이유" in info_df.columns:
            st.write(f"- 품목 선정이유: **{info_df.iloc[0]['선정이유']}**")

# ---------------------------------
# 3. 체인별 비교
# ---------------------------------
elif menu == "체인별 비교":
    st.header("🔍 체인별 비교")

    st.dataframe(compare, use_container_width=True, height=300)

    if "체인구분" in compare.columns and "평균_최종위험점수" in compare.columns:
        fig1 = px.bar(
            compare,
            x="체인구분",
            y="평균_최종위험점수",
            color="체인구분",
            text_auto=".2f",
            title="체인별 평균 최종위험점수"
        )
        fig1.update_layout(height=420)
        st.plotly_chart(fig1, use_container_width=True)

    if "체인구분" in compare.columns and "주력_상위1국의존도 (%)" in compare.columns:
        fig2 = px.bar(
            compare,
            x="체인구분",
            y="주력_상위1국의존도 (%)",
            color="체인구분",
            text_auto=".2f",
            title="체인별 상위1국 의존도"
        )
        fig2.update_layout(height=420)
        st.plotly_chart(fig2, use_container_width=True)

# ---------------------------------
# 4. 방법론 설명
# ---------------------------------
elif menu == "방법론 설명":
    st.header("📘 방법론 설명")

    st.markdown("""
### 1) 국가위험 기초점수
- `RISK_MASTER` 단계에서 고정가중치를 적용
- 2021~2024: `0.6 × UCDP + 0.25 × WGI + 0.15 × GPI`
- 2025: `0.6 × ACLED + 0.25 × WGI + 0.15 × GPI`

### 2) 국가별 보정단계
- `COUNTRY_MONTHLY` 단계에서 규칙기반 보정 적용
- 공급국 집중, 지역권 집중, HHI, 상위공급국 여부, FTA 여부 반영

### 3) 최종 공급망 위험 통합
- `PANEL_MONTHLY`, `ALERT_RESULT` 단계에서 가격·수급·물류·정책이벤트 리스크 통합

### 4) 엔트로피 가중치
- `ENTROPY_WEIGHT` 시트에서 확인 가능
- 단일 변수 카테고리는 내부 가중치가 100%로 보일 수 있음
    """)

    st.markdown("### 엔트로피 가중치 시트")
    st.dataframe(entropy, use_container_width=True, height=420)

# ---------------------------------
# 5. 시트 미리보기
# ---------------------------------
elif menu == "시트 미리보기":
    st.header("🧾 시트 미리보기")

    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    selected_sheet = st.selectbox("시트 선택", sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    st.write(f"선택 시트: **{selected_sheet}**")
    st.dataframe(df.head(30), use_container_width=True, height=500)
