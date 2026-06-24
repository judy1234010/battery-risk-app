import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="배터리 공급망 리스크 조기경보 플랫폼",
    page_icon="🔋",
    layout="wide"
)

# =========================================================
# 공통 함수
# =========================================================
def safe_read_excel(file_bytes, sheet_name):
    try:
        return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
    except Exception:
        return None

def clean_columns(df):
    if df is None:
        return None
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def fmt_num(x, digits=2):
    try:
        if pd.isna(x):
            return "-"
        return f"{float(x):,.{digits}f}"
    except Exception:
        return "-"

def fmt_pct(x, digits=2, ratio=False):
    try:
        if pd.isna(x):
            return "-"
        v = float(x)
        if ratio:
            v *= 100
        return f"{v:,.{digits}f}%"
    except Exception:
        return "-"

def traffic_emoji(level):
    s = str(level)
    if "심각" in s or "매우높음" in s:
        return "🔴"
    if "경계" in s or "높음" in s:
        return "🟠"
    if "주의" in s or "보통" in s:
        return "🟡"
    if "정상" in s or "낮음" in s:
        return "🟢"
    return "⚪"

def safe_columns(df, cols):
    return [c for c in cols if c in df.columns]

def get_latest_index(options):
    return max(len(options) - 1, 0)

def parse_ym_value(v):
    """
    이 파일의 연월 표기를 통일합니다.
    실제 확인 결과 10월은 여러 시트에서 '2021.1', '2022.1'처럼 저장되어 있습니다.
    따라서 아래 규칙을 적용합니다.

    - 2021.01 ~ 2021.09 -> 2021-01 ~ 2021-09
    - 2021.1            -> 2021-10
    - 2021.11           -> 2021-11
    - 2021.12           -> 2021-12
    """
    if pd.isna(v):
        return None, None

    s = str(v).strip()

    if "." in s:
        left, right = s.split(".", 1)
        left = left.strip()
        right = right.strip()

        if left.isdigit() and len(left) == 4:
            year = int(left)

            if right == "1":
                month = 10
                return year * 100 + month, f"{year}-{month:02d}"

            if right in ["11", "12"]:
                month = int(right)
                return year * 100 + month, f"{year}-{month:02d}"

            if len(right) == 2 and right.isdigit() and right.startswith("0"):
                month = int(right)
                if 1 <= month <= 9:
                    return year * 100 + month, f"{year}-{month:02d}"

            if right.isdigit():
                month = int(right)
                if 1 <= month <= 12:
                    return year * 100 + month, f"{year}-{month:02d}"

    digits = "".join(ch for ch in s if ch.isdigit())

    if len(digits) == 6:
        year = int(digits[:4])
        month = int(digits[4:])
        if 1 <= month <= 12:
            return year * 100 + month, f"{year}-{month:02d}"

    if len(digits) == 5:
        year = int(digits[:4])
        month = int(digits[4:])
        if 1 <= month <= 9:
            return year * 100 + month, f"{year}-{month:02d}"

    if len(digits) == 4:
        year = int(digits)
        return year * 100, f"{year}"

    return None, s

def add_ym_display(df):
    if df is None or "연월" not in df.columns:
        return df
    temp = df.copy()
    parsed = temp["연월"].apply(parse_ym_value)
    temp["연월_sort"] = parsed.apply(lambda x: x[0] if x else None)
    temp["연월_표시"] = parsed.apply(lambda x: x[1] if x else None)
    return temp

def coerce_numeric(df, cols):
    if df is None:
        return df
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# =========================================================
# 데이터 로드
# =========================================================
@st.cache_data
def load_all_data(file_bytes):
    data = {
        "ITEM_INFO": clean_columns(safe_read_excel(file_bytes, "ITEM_INFO")),
        "RISK_MASTER": clean_columns(safe_read_excel(file_bytes, "RISK_MASTER")),
        "COUNTRY_MONTHLY": clean_columns(safe_read_excel(file_bytes, "COUNTRY_MONTHLY")),
        "TPU_INDEX": clean_columns(safe_read_excel(file_bytes, "TPU_INDEX")),
        "MARKET_INDEX": clean_columns(safe_read_excel(file_bytes, "MARKET_INDEX")),
        "GSCPI_INDEX": clean_columns(safe_read_excel(file_bytes, "GSCPI_INDEX")),
        "HS_MONTHLY_SUMMARY": clean_columns(safe_read_excel(file_bytes, "HS_MONTHLY_SUMMARY")),
        "PANEL_MONTHLY": clean_columns(safe_read_excel(file_bytes, "PANEL_MONTHLY")),
        "ALERT_RESULT": clean_columns(safe_read_excel(file_bytes, "ALERT_RESULT")),
        "체인별 비교표": clean_columns(safe_read_excel(file_bytes, "체인별 비교표")),
        "ENTROPY_WEIGHT": clean_columns(safe_read_excel(file_bytes, "ENTROPY_WEIGHT")),
        "NOMALIZATION_CHECK": clean_columns(safe_read_excel(file_bytes, "NOMALIZATION_CHECK")),
    }

    for key in ["COUNTRY_MONTHLY", "TPU_INDEX", "MARKET_INDEX", "GSCPI_INDEX", "HS_MONTHLY_SUMMARY", "PANEL_MONTHLY", "ALERT_RESULT", "NOMALIZATION_CHECK"]:
        data[key] = add_ym_display(data[key])

    return data

# =========================================================
# 헤더
# =========================================================
st.title("🔋 배터리 공급망 리스크 조기경보 플랫폼")
st.caption("수입구조, 국가위험, 가격·물류·정책충격을 통합하여 공급망 리스크를 진단하는 공모전 최종 발표 및 제출용 플랫폼입니다.")

uploaded_file = st.file_uploader("최종 엑셀 파일(.xlsx)을 업로드해 주십시오.", type=["xlsx"])

if uploaded_file is None:
    st.info("플랫폼을 시작하려면 최종 엑셀 파일을 업로드해 주십시오.")
    st.stop()

file_bytes = uploaded_file.getvalue()
data = load_all_data(file_bytes)

item_info = data["ITEM_INFO"]
risk_master = data["RISK_MASTER"]
country = data["COUNTRY_MONTHLY"]
tpu = data["TPU_INDEX"]
market = data["MARKET_INDEX"]
gscpi = data["GSCPI_INDEX"]
hs_summary = data["HS_MONTHLY_SUMMARY"]
panel = data["PANEL_MONTHLY"]
alert = data["ALERT_RESULT"]
compare = data["체인별 비교표"]
entropy = data["ENTROPY_WEIGHT"]
norm_check = data["NOMALIZATION_CHECK"]

required_sheets = {
    "COUNTRY_MONTHLY": country,
    "TPU_INDEX": tpu,
    "MARKET_INDEX": market,
    "GSCPI_INDEX": gscpi,
    "HS_MONTHLY_SUMMARY": hs_summary,
    "PANEL_MONTHLY": panel,
    "ALERT_RESULT": alert,
    "체인별 비교표": compare,
    "ENTROPY_WEIGHT": entropy,
}

missing = [k for k, v in required_sheets.items() if v is None]
if missing:
    st.error(f"필수 시트를 찾지 못했습니다: {', '.join(missing)}")
    st.stop()

# =========================================================
# 전처리
# =========================================================
for df in [country, hs_summary, panel, alert]:
    if df is not None and "체인구분" in df.columns:
        df["체인구분"] = df["체인구분"].astype(str)

country = coerce_numeric(country, ["국가별수입비중", "지역권별수입비중", "기본평가점수", "총보정점수", "최종보정점수"])
panel = coerce_numeric(panel, [
    "총수입금액", "총수입물량", "평균수입단가", "상위1국의존도", "상위3국집중도", "HHI", "수입국수", "CV",
    "국가보정합계", "환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화", "GSCPI", "GSCPI_Norm",
    "원천_TPU_INDEX", "TPU_Norm", "환율정규화_5pt", "납가격정규화_5pt", "리튬가격정규화_5pt", "니켈가격정규화_5pt",
    "HHI_5pt", "국가보정합계_5pt", "GSCPI_Norm_5pt", "TPU_Norm_5pt", "가격리스크점수", "수급리스크점수",
    "물류리스크점수", "정책이벤트리스크점수", "최종위험점수_raw", "최종위험점수"
])
alert = coerce_numeric(alert, ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수", "fta_ratio", "대체조달가능성_점수"])
market = coerce_numeric(market, ["환율", "월평균납가격", "월평균리튬가격", "월평균니켈가격", "환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화",
                                 "환율_전월비", "납가격_전월비", "리튬가격_전월비", "니켈가격_전월비", "환율_3개월변동", "납가격_3개월변동", "리튬가격_3개월변동", "니켈가격_3개월변동"])
gscpi = coerce_numeric(gscpi, ["GSCPI", "GSCPI_NORM"])
tpu = coerce_numeric(tpu, ["원천_TPU_INDEX", "이벤트보정"])
if norm_check is not None:
    norm_check = coerce_numeric(norm_check, [
        "환율", "월평균납가격", "월평균리튬가격", "월평균니켈가격",
        "환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화",
        "GSCPI", "GSCPI_NORM", "원천_TPU_INDEX", "이벤트보정"
    ])

chains = sorted(panel["체인구분"].dropna().unique().tolist())
month_labels = panel[["연월_sort", "연월_표시"]].dropna().drop_duplicates().sort_values("연월_sort")
month_options = month_labels["연월_표시"].tolist()

st.sidebar.title("메뉴")
menu = st.sidebar.radio(
    "페이지 선택",
    [
        "1. 조기경보 대시보드",
        "2. 품목·국가 취약성 분석",
        "3. 충격 원인 추적",
        "4. 국가위험 근거 보기",
        "5. 기업 대응 시뮬레이터",
        "6. 방법론 / 가중치 설명",
    ]
)

# =========================================================
# 1. 조기경보 대시보드
# =========================================================
if menu == "1. 조기경보 대시보드":
    st.header("📊 공급망 조기경보 종합 현황")
    st.markdown("**페이지 활용 목적:** 특정 체인의 현재 위험 수준과 공급집중 구조를 종합적으로 진단하기 위한 화면입니다.")

    c1, c2 = st.columns(2)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", chains)
    with c2:
        selected_month_label = st.selectbox("연월 선택", month_options, index=get_latest_index(month_options))

    panel_row = panel[(panel["체인구분"] == selected_chain) & (panel["연월_표시"] == selected_month_label)]
    alert_row = alert[(alert["체인구분"] == selected_chain) & (alert["연월_표시"] == selected_month_label)]

    if panel_row.empty or alert_row.empty:
        st.warning("선택한 조건의 데이터가 없습니다.")
        st.stop()

    p = panel_row.iloc[0]
    a = alert_row.iloc[0]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("최종위험점수", fmt_num(a.get("최종위험점수")))
    m2.metric("최종경보등급", f"{traffic_emoji(a.get('최종경보등급'))} {a.get('최종경보등급', '-')}")
    m3.metric("대체조달가능성", str(a.get("대체조달가능성", "-")))
    m4.metric("우선관리대상", str(a.get("상대적_우선관리대상", "-")))

    m5, m6, m7, m8 = st.columns(4)
    m5.metric("FTA 활용비중", fmt_pct(a.get("fta_ratio"), ratio=False))
    m6.metric("상위1국의존도", fmt_pct(p.get("상위1국의존도"), ratio=False))
    m7.metric("상위3국집중도", fmt_pct(p.get("상위3국집중도"), ratio=False))
    m8.metric("HHI", fmt_num(p.get("HHI")))

    risk_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    top_risk = "-"
    if all(col in alert.columns for col in risk_cols):
        risk_df = pd.DataFrame({
            "리스크유형": ["가격", "수급", "물류", "정책이벤트"],
            "점수": [a.get("가격리스크점수"), a.get("수급리스크점수"), a.get("물류리스크점수"), a.get("정책이벤트리스크점수")]
        })
        fig = px.bar(risk_df, x="리스크유형", y="점수", color="리스크유형", text_auto=".2f",
                     title=f"{selected_chain} / {selected_month_label} 리스크 유형별 점수")
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)
        top_risk = risk_df.sort_values("점수", ascending=False).iloc[0]["리스크유형"]

    st.markdown("### 월별 최종위험점수 추이")
    chain_ts = alert[alert["체인구분"] == selected_chain].copy().sort_values("연월_sort")
    fig2 = px.line(chain_ts, x="연월_표시", y="최종위험점수", markers=True,
                   title=f"{selected_chain} 월별 최종위험점수 추이")
    fig2.update_layout(height=420, xaxis_title="연월", yaxis_title="최종위험점수")
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 분석 의의")
    st.write(f"- {selected_month_label} 기준 {selected_chain}의 최종위험점수는 **{fmt_num(a.get('최종위험점수'))}점**이며, 경보등급은 **{a.get('최종경보등급', '-')} 단계**로 판단됩니다.")
    st.write(f"- 상위 1개국 의존도는 **{fmt_pct(p.get('상위1국의존도'), ratio=False)}**, 상위 3개국 집중도는 **{fmt_pct(p.get('상위3국집중도'), ratio=False)}**, HHI는 **{fmt_num(p.get('HHI'))}**로 확인되어 공급집중 구조의 강도를 함께 설명합니다.")
    st.write(f"- 대체조달가능성은 **{a.get('대체조달가능성', '-')}**, FTA 활용비중은 **{fmt_pct(a.get('fta_ratio'), ratio=False)}**로 나타납니다.")
    st.write(f"- 해당 시점에서 상대적으로 가장 크게 작용한 요인은 **{top_risk} 리스크**로 해석됩니다.")
    st.write(f"- 보정사유는 **{a.get('보정사유', '-')}**이며, 비고는 **{a.get('비고', '-')}**입니다.")

# =========================================================
# 2. 품목·국가 취약성 분석
# =========================================================
elif menu == "2. 품목·국가 취약성 분석":
    st.header("🌍 품목별·국가별 취약성 진단")
    st.markdown("**페이지 활용 목적:** 특정 품목이 어느 국가에 집중되어 있으며, 어떤 보정요인으로 인해 취약성이 높아졌는지 확인하기 위한 화면입니다.")

    c1, c2, c3 = st.columns(3)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", sorted(country["체인구분"].dropna().unique().tolist()))

    chain_months = country[country["체인구분"] == selected_chain][["연월_sort", "연월_표시"]].dropna().drop_duplicates().sort_values("연월_sort")
    month_opts = chain_months["연월_표시"].tolist()

    with c2:
        selected_month_label = st.selectbox("연월 선택", month_opts, index=get_latest_index(month_opts))

    hs_candidates = sorted(
        country[
            (country["체인구분"] == selected_chain) &
            (country["연월_표시"] == selected_month_label)
        ]["HS코드"].astype(str).str.replace(".0", "", regex=False).dropna().unique().tolist()
    )

    with c3:
        selected_hs = st.selectbox("HS코드 선택", hs_candidates)

    detail_df = country[
        (country["체인구분"] == selected_chain) &
        (country["연월_표시"] == selected_month_label) &
        (country["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs)
    ].copy()

    summary_df = hs_summary[
        (hs_summary["체인구분"] == selected_chain) &
        (hs_summary["연월_표시"] == selected_month_label) &
        (hs_summary["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs)
    ].copy()

    if detail_df.empty:
        st.warning("선택한 조건의 데이터가 없습니다.")
        st.stop()

    detail_df["국가별수입비중"] = pd.to_numeric(detail_df["국가별수입비중"], errors="coerce")
    detail_df = detail_df.sort_values("국가별수입비중", ascending=False)

    if not summary_df.empty:
        h = summary_df.iloc[0]
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("품목명", str(h.get("품목명", "-")))
        s2.metric("총수입금액", fmt_num(h.get("총수입금액"), 0))
        s3.metric("상위공급국", str(h.get("상위공급국", "-")))
        s4.metric("상위공급국비중", fmt_pct(h.get("상위공급국비중"), ratio=True))

        s5, s6, s7, s8 = st.columns(4)
        s5.metric("수입국수", fmt_num(h.get("수입국수"), 0))
        s6.metric("평균수입단가", fmt_num(h.get("평균수입단가")))
        s7.metric("전월대비 수입금액증감률", fmt_pct(h.get("전월대비수입금액증감률"), ratio=True))
        s8.metric("전년동월대비 수입금액증감률", fmt_pct(h.get("전년동월대비수입금액증감률"), ratio=True))

    st.markdown("### 국가별 취약성 상세")
    show_cols = safe_columns(detail_df, ["국가명", "지역권", "FTA여부", "국가별수입비중", "지역권별수입비중",
                                         "상위공급국여부", "기본평가점수", "총보정점수", "최종보정점수", "최종판정", "비고"])
    st.dataframe(detail_df[show_cols], use_container_width=True, height=420)

    if "국가명" in detail_df.columns and "국가별수입비중" in detail_df.columns:
        fig = px.bar(detail_df.head(10), x="국가명", y="국가별수입비중",
                     color="지역권" if "지역권" in detail_df.columns else None,
                     text_auto=".2f", title="상위 국가별 수입비중(%)")
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    top_row = detail_df.iloc[0]
    st.markdown("### 분석 의의")
    st.write(f"- {selected_month_label} 기준 HS {selected_hs} 품목은 **{top_row.get('국가명', '-')}**에 가장 크게 의존하고 있으며, 해당 국가 수입비중은 **{fmt_pct(top_row.get('국가별수입비중'), ratio=False)}**입니다.")
    if "FTA여부" in detail_df.columns:
        fta_n = detail_df[detail_df["FTA여부"] == "N"]
        st.write(f"- FTA 미체결 또는 비활용 국가 건수는 **{len(fta_n)}건**으로 확인됩니다.")
    if "최종판정" in detail_df.columns:
        high_risk = detail_df[detail_df["최종판정"].isin(["높음", "매우높음"])]
        st.write(f"- 최종판정 기준 고위험 국가 수는 **{len(high_risk)}건**입니다.")
    if item_info is not None and "HS코드" in item_info.columns:
        info_df = item_info[item_info["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs]
        if not info_df.empty and "선정이유" in info_df.columns:
            st.write(f"- 품목 선정이유는 다음과 같습니다: **{info_df.iloc[0]['선정이유']}**")

# =========================================================
# 3. 충격 원인 추적
# =========================================================
elif menu == "3. 충격 원인 추적":
    st.header("📈 충격 원인 추적 및 선제대응 시사점")
    st.markdown("**페이지 활용 목적:** 특정 시점의 위험 상승이 가격, 물류, 정책이벤트 중 어떤 요인에서 비롯되었는지 추적하기 위한 화면입니다.")

    c1, c2 = st.columns(2)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", chains)
    with c2:
        selected_month_label = st.selectbox("기준 연월 선택", month_options, index=get_latest_index(month_options))

    panel_row = panel[(panel["체인구분"] == selected_chain) & (panel["연월_표시"] == selected_month_label)]
    alert_row = alert[(alert["체인구분"] == selected_chain) & (alert["연월_표시"] == selected_month_label)]

    if panel_row.empty or alert_row.empty:
        st.warning("선택한 조건의 데이터가 없습니다.")
        st.stop()

    p = panel_row.iloc[0]
    a = alert_row.iloc[0]

    shock_df = norm_check if norm_check is not None else market
    if shock_df is not None:
        shock_df = shock_df.sort_values("연월_sort").copy()

    st.markdown("### 1) 가격충격 추이")
    if shock_df is not None:
        price_norm_cols = [c for c in ["환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화"] if c in shock_df.columns]
        if price_norm_cols:
            fig1 = px.line(shock_df, x="연월_표시", y=price_norm_cols,
                           title="가격충격 정규화 지표 추이(원 파일 산식 기준)")
            fig1.update_layout(height=420, xaxis_title="연월", yaxis_title="정규화 지수")
            st.plotly_chart(fig1, use_container_width=True)

        sel = shock_df[shock_df["연월_표시"] == selected_month_label]
        if not sel.empty:
            sr = sel.iloc[0]
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("환율 정규화", fmt_num(sr.get("환율정규화"), 4))
            d2.metric("납가격 정규화", fmt_num(sr.get("납가격정규화"), 4))
            d3.metric("리튬가격 정규화", fmt_num(sr.get("리튬가격정규화"), 4))
            d4.metric("니켈가격 정규화", fmt_num(sr.get("니켈가격정규화"), 4))

    price_5pt = pd.DataFrame({
        "항목": ["환율", "납가격", "리튬가격", "니켈가격"],
        "5점 환산값": [
            p.get("환율정규화_5pt"),
            p.get("납가격정규화_5pt"),
            p.get("리튬가격정규화_5pt"),
            p.get("니켈가격정규화_5pt"),
        ]
    })
    fig1b = px.bar(price_5pt, x="항목", y="5점 환산값", color="항목", text_auto=".2f",
                   title=f"{selected_month_label} 가격충격 세부 기여도(5점 환산)")
    fig1b.update_layout(height=380)
    st.plotly_chart(fig1b, use_container_width=True)

    st.markdown("### 2) 정책이벤트 추이")
    if tpu is not None:
        tpu_plot = tpu.sort_values("연월_sort").copy()
        if "원천_TPU_INDEX" in tpu_plot.columns:
            fig2 = px.line(tpu_plot, x="연월_표시", y="원천_TPU_INDEX", markers=True,
                           title="TPU(통상정책 불확실성) 추이")
            fig2.update_layout(height=380, xaxis_title="연월", yaxis_title="TPU 지수")
            st.plotly_chart(fig2, use_container_width=True)

        tpu_sel = tpu_plot[tpu_plot["연월_표시"] == selected_month_label]
        if not tpu_sel.empty:
            tr = tpu_sel.iloc[0]
            t1, t2 = st.columns(2)
            t1.metric("TPU 지수", fmt_num(tr.get("원천_TPU_INDEX"), 2))
            t2.metric("이벤트 보정", fmt_num(tr.get("이벤트보정"), 2))
            st.write(f"**서사 배경:** {tr.get('서사배경', '-')}")

    st.markdown("### 3) 물류충격 추이")
    if gscpi is not None:
        g_plot = gscpi.sort_values("연월_sort").copy()
        if "GSCPI_NORM" in g_plot.columns:
            fig3 = px.line(g_plot, x="연월_표시", y="GSCPI_NORM", markers=True,
                           title="GSCPI 정규화 추이")
            fig3.update_layout(height=380, xaxis_title="연월", yaxis_title="정규화 지수")
            st.plotly_chart(fig3, use_container_width=True)

        gs = g_plot[g_plot["연월_표시"] == selected_month_label]
        if not gs.empty:
            gr = gs.iloc[0]
            g1, g2 = st.columns(2)
            g1.metric("GSCPI", fmt_num(gr.get("GSCPI"), 2))
            g2.metric("GSCPI 정규화", fmt_num(gr.get("GSCPI_NORM"), 2))

    st.markdown("### 4) 리스크 기여도 종합")
    risk_df = pd.DataFrame({
        "유형": ["가격", "수급", "물류", "정책이벤트"],
        "점수": [a.get("가격리스크점수"), a.get("수급리스크점수"), a.get("물류리스크점수"), a.get("정책이벤트리스크점수")]
    }).sort_values("점수", ascending=False)

    fig4 = px.bar(risk_df, x="유형", y="점수", color="유형", text_auto=".2f",
                  title=f"{selected_month_label} 최종 리스크 기여도")
    fig4.update_layout(height=380)
    st.plotly_chart(fig4, use_container_width=True)

    dominant = risk_df.iloc[0]["유형"]

    st.markdown("### 분석 의의")
    st.write(f"- {selected_month_label} 기준 {selected_chain}의 최종위험점수는 **{fmt_num(a.get('최종위험점수'))}점**이며, 핵심 상승요인은 **{dominant} 리스크**로 해석됩니다.")
    st.write(f"- 가격충격 세부항목의 5점 환산 결과는 환율 **{fmt_num(p.get('환율정규화_5pt'))}**, 납가격 **{fmt_num(p.get('납가격정규화_5pt'))}**, 리튬가격 **{fmt_num(p.get('리튬가격정규화_5pt'))}**, 니켈가격 **{fmt_num(p.get('니켈가격정규화_5pt'))}**입니다.")
    st.write(f"- 물류충격은 GSCPI 정규화값 **{fmt_num(p.get('GSCPI_Norm'))}** 및 5점 환산값 **{fmt_num(p.get('GSCPI_Norm_5pt'))}**로 반영되며, 정책이벤트 충격은 TPU 지수 및 이벤트 보정에 기반한 **{fmt_num(p.get('TPU_Norm_5pt'))}점 수준**으로 반영됩니다.")

    if dominant == "정책이벤트":
        st.write("- 이는 통상정책 변화 또는 지정학적 이벤트가 공급망에 선행 압력으로 작용했음을 시사하며, 향후 2~3개월 내 대체국 확보와 조달선 재점검이 필요함을 의미합니다.")
    elif dominant == "물류":
        st.write("- 이는 글로벌 운송 병목과 물류 압력이 실질적 위험요인으로 작동했음을 시사하며, 향후 1~2개월 내 선복, 운송, 통관 모니터링 강화가 필요함을 의미합니다.")
    elif dominant == "가격":
        st.write("- 이는 환율 및 원자재 가격 변동이 수입단가와 비용 구조에 직접적인 압력을 가하고 있음을 시사하며, 향후 1~2개월 내 선매입, 계약단가 조정, 가격헤지 검토가 필요함을 의미합니다.")
    else:
        st.write("- 이는 공급집중 구조와 국가별 보정요인이 복합적으로 작용했음을 의미하며, 향후 2~3개월 내 수입선 다변화와 FTA 활용 확대 전략이 필요합니다.")

# =========================================================
# 4. 국가위험 근거 보기
# =========================================================
elif menu == "4. 국가위험 근거 보기":
    st.header("🛡️ 국가위험 점수의 구조적 근거")
    st.markdown("**페이지 활용 목적:** 국가위험 점수가 임의의 판단값이 아니라, 분쟁·거버넌스·평화지표를 결합한 구조적 지표임을 설명하기 위한 화면입니다.")

    if risk_master is None:
        st.warning("RISK_MASTER 시트를 찾지 못했습니다.")
        st.stop()

    country_options = sorted(risk_master["Country_std"].dropna().unique().tolist())
    year_options = sorted(risk_master["Year"].dropna().unique().tolist())

    c1, c2 = st.columns(2)
    with c1:
        selected_country = st.selectbox("국가 선택", country_options)
    with c2:
        selected_year = st.selectbox("연도 선택", year_options, index=get_latest_index(year_options))

    rr = risk_master[(risk_master["Country_std"] == selected_country) & (risk_master["Year"] == selected_year)]
    if rr.empty:
        st.warning("선택한 국가/연도 데이터가 없습니다.")
        st.stop()

    r = rr.iloc[0]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Composite Risk Score", fmt_num(r.get("Composite_Risk_Score")))
    k2.metric("Risk Flag", f"{traffic_emoji(r.get('Risk_Flag'))} {r.get('Risk_Flag', '-')}")
    k3.metric("WGI Risk Score", fmt_num(r.get("WGI_Risk_Score_Used")))
    k4.metric("GPI Risk Score", fmt_num(r.get("GPI_Risk_Score")))

    k5, k6, k7 = st.columns(3)
    k5.metric("UCDP Risk Score", fmt_num(r.get("UCDP_Risk_Score")))
    k6.metric("ACLED Risk Score", fmt_num(r.get("ACLED_Risk_Score")))
    k7.metric("Composite Method", str(r.get("Composite_Method", "-")))

    st.markdown("### 분석 의의")
    st.write(f"- {selected_country}의 {selected_year}년 종합위험점수는 **{fmt_num(r.get('Composite_Risk_Score'))}점**입니다.")
    st.write(f"- 최종 위험 판정은 **{r.get('Risk_Flag', '-')}**이며, 합성방식은 **{r.get('Composite_Method', '-')}**입니다.")
    if pd.notna(r.get("ACLED_Risk_Score")):
        st.write("- 최근 연도에는 ACLED 기반 분쟁위험이 반영되어 최근성 및 현시점 반영력이 강화되었습니다.")
    else:
        st.write("- 과거 연도에는 UCDP 기반 분쟁위험이 활용되어 연도 간 비교 가능성이 유지됩니다.")
    st.write("- 따라서 본 지표는 직관적 판단이 아니라, 설명 가능한 구조적 위험평가 체계로 제시될 수 있습니다.")

# =========================================================
# 5. 기업 대응 시뮬레이터
# =========================================================
elif menu == "5. 기업 대응 시뮬레이터":
    st.header("🏭 기업 대응 전략 시뮬레이터")
    st.markdown("**페이지 활용 목적:** 기업이 수입선 다변화, FTA 활용, 외부충격 대응 전략을 가정하여 위험 완화 방향을 시뮬레이션하기 위한 화면입니다.")

    c1, c2, c3 = st.columns(3)
    with c1:
        current_top1 = st.slider("현재 상위1국의존도(%)", 0.0, 100.0, 70.0, 0.5)
    with c2:
        current_hhi = st.slider("현재 HHI", 0.0, 10000.0, 5000.0, 10.0)
    with c3:
        current_adj = st.slider("현재 국가보정합계", 0.0, 100.0, 40.0, 0.5)

    c4, c5, c6 = st.columns(3)
    with c4:
        diversify = st.slider("다변화 수준(상위1국 비중 감소폭)", 0.0, 50.0, 10.0, 0.5)
    with c5:
        fta_change = st.selectbox("FTA 활용 전략", ["변화 없음", "FTA 신규 활용"])
    with c6:
        shock_type = st.selectbox("예상 외부충격", ["없음", "정책이벤트 충격", "물류충격", "가격충격"])

    new_top1 = max(current_top1 - diversify, 0)
    new_hhi = max(current_hhi - diversify * 55, 0)
    new_adj = max(current_adj - diversify * 0.35, 0)

    if fta_change == "FTA 신규 활용":
        new_adj = max(new_adj - 5, 0)

    shock_add = 0
    if shock_type == "정책이벤트 충격":
        shock_add = 8
    elif shock_type == "물류충격":
        shock_add = 6
    elif shock_type == "가격충격":
        shock_add = 5

    before_score = current_adj + current_top1 * 0.2 + current_hhi / 1000
    after_score = new_adj + new_top1 * 0.2 + new_hhi / 1000 + shock_add

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("변경 후 상위1국의존도", f"{new_top1:.2f}%", f"{new_top1 - current_top1:.2f}")
    r2.metric("변경 후 HHI", f"{new_hhi:.2f}", f"{new_hhi - current_hhi:.2f}")
    r3.metric("변경 후 국가보정합계", f"{new_adj:.2f}", f"{new_adj - current_adj:.2f}")
    r4.metric("시뮬레이션 위험지수", f"{after_score:.2f}", f"{after_score - before_score:.2f}")

    st.markdown("### 전략적 시사점")
    if diversify >= 10:
        st.write("- 수입선 다변화 폭이 유의미한 수준이므로, 공급집중 리스크 완화 효과를 기대할 수 있습니다.")
    else:
        st.write("- 다변화 폭이 제한적이므로, 구조적 위험 완화 효과는 크지 않을 가능성이 있습니다.")
    if fta_change == "FTA 신규 활용":
        st.write("- FTA 신규 활용은 통상비용 완충과 국가보정 리스크 완화에 기여할 수 있습니다.")
    if shock_type == "정책이벤트 충격":
        st.write("- 정책이벤트 충격이 발생하는 경우, 향후 2~3개월 내 조달선 재편과 재고전략 점검이 필요합니다.")
    elif shock_type == "물류충격":
        st.write("- 물류충격이 발생하는 경우, 향후 1~2개월 내 선복 확보와 통관 모니터링 강화가 중요합니다.")
    elif shock_type == "가격충격":
        st.write("- 가격충격이 발생하는 경우, 단가 재협상, 선매입, 가격헤지 검토가 필요합니다.")
    if new_top1 < 50:
        st.write("- 상위1국의존도를 50% 미만으로 낮출 경우, 구조적 취약성 완화 효과를 보다 명확히 설명할 수 있습니다.")

# =========================================================
# 6. 방법론 / 가중치 설명
# =========================================================
elif menu == "6. 방법론 / 가중치 설명":
    st.header("📘 방법론·가중치·재현가능성 설명")
    st.markdown("**페이지 활용 목적:** 본 모형이 어떤 지표를 어떤 방식으로 결합하여 최종위험점수를 도출했는지, 그리고 왜 설명 가능하고 재현 가능한 구조인지 제시하기 위한 화면입니다.")

    st.markdown("""
### 1) 모형 설계 원리
본 플랫폼은 단일 지표에 의존하지 않고, **수입구조 정보**, **국가위험 정보**, **시장가격 충격**, **물류 충격**, **정책이벤트 충격**을 결합하는 다층 구조로 설계되었습니다.  
따라서 특정 월의 위험도는 단순 수입금액 변화가 아니라, 공급집중 구조와 외생충격을 함께 반영한 결과입니다.

### 2) 국가위험 기초점수 산출 방식
`RISK_MASTER` 단계에서는 **고정가중치 방식**을 적용하였습니다.

- **2021~2024년**
  - `0.60 × UCDP_Risk_Score`
  - `0.25 × WGI_Risk_Score_Used`
  - `0.15 × GPI_Risk_Score`

- **2025년**
  - `0.60 × ACLED_Risk_Score`
  - `0.25 × WGI_Risk_Score_Used`
  - `0.15 × GPI_Risk_Score`

이는 분쟁위험을 핵심축으로 두되, 거버넌스와 평화 수준을 함께 반영하여 **구조적 국가위험**을 측정하기 위한 설계입니다.

### 3) 국가별 보정단계의 논리
`COUNTRY_MONTHLY` 단계에서는 국가별 기본위험점수에 더하여, 실제 수입구조에서 발생하는 취약성을 **규칙기반 보정식**으로 반영하였습니다.

주요 보정항목은 다음과 같습니다.

- 공급국 집중도 보정
- 지역권 집중도 보정
- HHI 기반 집중도 보정
- 상위공급국 여부 보정
- FTA 여부 보정

즉 동일한 국가라도, 해당 국가에 대한 실제 의존도가 높을수록 최종 취약성은 더 크게 반영됩니다.

### 4) 외생충격 지표의 반영 방식
`MARKET_INDEX`, `GSCPI_INDEX`, `TPU_INDEX` 단계에서는 환율, 원자재 가격, 글로벌 공급망 압력, 통상정책 불확실성을 월별 시계열로 정규화하였습니다.  
이 외생충격 지표는 `PANEL_MONTHLY` 단계에서 가격·물류·정책이벤트 리스크의 세부 입력값으로 반영됩니다.

### 5) 최종 통합단계의 가중치 방식
`PANEL_MONTHLY` 및 `ALERT_RESULT` 단계에서는 가격, 수급, 물류, 정책이벤트 리스크를 통합하기 위해 **엔트로피 가중치 방식**을 활용하였습니다.

엔트로피 방식은 데이터의 분산도와 정보량을 기준으로 가중치를 도출하므로, 연구자가 임의로 중요도를 부여하는 방식보다 **객관성과 재현가능성**이 높습니다.

### 6) 물류 / 정책이벤트가 100%로 표시되는 이유
일부 카테고리는 내부 구성변수가 1개뿐이므로, 카테고리 내부 가중치가 100%로 표시됩니다.  
이는 전체 모형이 단일 변수라는 뜻이 아니라, **해당 하위 카테고리 내부에서 유일한 변수**라는 의미입니다.

### 7) 설명가능성과 재현가능성의 의의
본 모형은 다음 세 가지 측면에서 설명가능성과 재현가능성을 갖습니다.

1. **입력 지표가 명확합니다.**  
   어떤 시트의 어떤 컬럼이 어떤 위험범주에 반영되는지 추적 가능합니다.

2. **결합 방식이 명확합니다.**  
   고정가중치, 규칙기반 보정, 엔트로피 가중치가 단계별로 구분되어 있습니다.

3. **동일 데이터에 대해 동일 결과가 재현됩니다.**  
   임의 판단이 아니라 수식 및 규칙 기반으로 점수가 산출되므로 반복가능성이 확보됩니다.

### 8) 정책 및 실무 활용성
따라서 본 플랫폼은 단순 시각화 도구를 넘어,

- 관세·통관 당국의 선제적 모니터링,
- 기업의 수입선 다변화 전략 수립,
- 공급망 이상징후 조기 탐지,
- 품목별 취약성 설명자료 작성

에 활용 가능한 실질적 의사결정 지원도구로 해석할 수 있습니다.
    """)

    st.markdown("### 엔트로피 가중치 현황")
    st.dataframe(entropy, use_container_width=True, height=360)

    st.markdown("### 체인별 비교 요약")
    st.dataframe(compare, use_container_width=True, height=240)

    st.markdown("### 종합 의의")
    st.write("- 본 모형은 구조적 국가위험, 실제 수입집중도, 외생 충격요인을 통합함으로써 공급망 위험을 다층적으로 설명할 수 있다는 점에서 의의가 있습니다.")
    st.write("- 또한 산출 논리와 가중치 체계가 명시되어 있어 설명가능성과 재현가능성이 확보된다는 점에서 공공정책 및 실무 활용 가치가 높습니다.")
