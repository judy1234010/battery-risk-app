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

def parse_ym_value(v):
    """
    다양한 연월 표현(2021.01, 202101, 20251, 2025.1 등)을
    (정렬키, 표시문자열 'YYYY-MM')로 변환
    """
    if pd.isna(v):
        return None, None

    s = str(v).strip()

    # 1) float처럼 들어온 경우: 2021.01 / 2025.1 / 2025.11
    if "." in s:
        left, right = s.split(".", 1)
        year = left.strip()

        if len(year) == 4 and year.isdigit():
            # 2021.01 -> 01
            # 2025.1  -> 10 또는 01이 아니라 실제 원본 의미가 애매할 수 있음
            # 데이터 특성상 2021.01 같은 형태가 많으므로 우선 소수부를 월로 해석
            month_part = right.strip()

            # 01, 02, 10, 11, 12 등 처리
            if month_part.isdigit():
                month = int(month_part)
                if 1 <= month <= 12:
                    label = f"{year}-{month:02d}"
                    sort_key = int(year) * 100 + month
                    return sort_key, label

    # 2) 순수 숫자형 문자열 처리: 202101 / 20251 / 202510
    digits = "".join(ch for ch in s if ch.isdigit())

    # YYYYMM
    if len(digits) == 6:
        year = int(digits[:4])
        month = int(digits[4:])
        if 1 <= month <= 12:
            return year * 100 + month, f"{year}-{month:02d}"

    # YYYYM (예: 20251 -> 2025-01, 20259 -> 2025-09)
    if len(digits) == 5:
        year = int(digits[:4])
        month = int(digits[4:])
        if 1 <= month <= 9:
            return year * 100 + month, f"{year}-{month:02d}"

    # 혹시 YYYY 형태만 온 경우
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

def fmt_num(x, digits=2):
    try:
        if pd.isna(x):
            return "-"
        return f"{float(x):,.{digits}f}"
    except Exception:
        return "-"

def fmt_pct(x, digits=2, already_percent=True):
    try:
        if pd.isna(x):
            return "-"
        v = float(x)
        if not already_percent:
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

# =========================================================
# 데이터 로드
# =========================================================
@st.cache_data
def load_all_data(file_bytes):
    data = {
        "ITEM_INFO": safe_read_excel(file_bytes, "ITEM_INFO"),
        "RISK_MASTER": safe_read_excel(file_bytes, "RISK_MASTER"),
        "COUNTRY_MONTHLY": safe_read_excel(file_bytes, "COUNTRY_MONTHLY"),
        "TPU_INDEX": safe_read_excel(file_bytes, "TPU_INDEX"),
        "MARKET_INDEX": safe_read_excel(file_bytes, "MARKET_INDEX"),
        "GSCPI_INDEX": safe_read_excel(file_bytes, "GSCPI_INDEX"),
        "HS_MONTHLY_SUMMARY": safe_read_excel(file_bytes, "HS_MONTHLY_SUMMARY"),
        "PANEL_MONTHLY": safe_read_excel(file_bytes, "PANEL_MONTHLY"),
        "ALERT_RESULT": safe_read_excel(file_bytes, "ALERT_RESULT"),
        "체인별 비교표": safe_read_excel(file_bytes, "체인별 비교표"),
        "ENTROPY_WEIGHT": safe_read_excel(file_bytes, "ENTROPY_WEIGHT"),
    }

    # 연월 표시 컬럼 부여
    for key in ["COUNTRY_MONTHLY", "TPU_INDEX", "MARKET_INDEX", "GSCPI_INDEX", "HS_MONTHLY_SUMMARY", "PANEL_MONTHLY", "ALERT_RESULT"]:
        data[key] = add_ym_display(data[key])

    # 문자열 정리
    for key in ["COUNTRY_MONTHLY", "HS_MONTHLY_SUMMARY", "PANEL_MONTHLY", "ALERT_RESULT", "체인별 비교표", "ENTROPY_WEIGHT", "ITEM_INFO", "RISK_MASTER"]:
        if data[key] is not None:
            data[key] = data[key].copy()

    return data

# =========================================================
# 헤더
# =========================================================
st.title("🔋 배터리 공급망 리스크 조기경보 플랫폼")
st.caption("수입통계·국가위험·시장충격·물류충격·정책이벤트를 결합한 발표용 의사결정 대시보드")

uploaded_file = st.file_uploader(
    "최종 엑셀 파일(.xlsx)을 업로드해줘",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info("발표용 대시보드를 시작하려면 엑셀 파일을 업로드해줘.")
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

required_sheets = {
    "COUNTRY_MONTHLY": country,
    "HS_MONTHLY_SUMMARY": hs_summary,
    "PANEL_MONTHLY": panel,
    "ALERT_RESULT": alert,
    "체인별 비교표": compare,
    "ENTROPY_WEIGHT": entropy,
}

missing = [k for k, v in required_sheets.items() if v is None]
if missing:
    st.error(f"필수 시트를 찾지 못했어: {', '.join(missing)}")
    st.stop()

# =========================================================
# 기본 전처리
# =========================================================
for df in [country, hs_summary, panel, alert]:
    if df is not None and "체인구분" in df.columns:
        df["체인구분"] = df["체인구분"].astype(str)

if country is not None and "국가별수입비중" in country.columns:
    country["국가별수입비중"] = pd.to_numeric(country["국가별수입비중"], errors="coerce")

if panel is not None:
    for col in ["상위1국의존도", "상위3국집중도", "HHI", "국가보정합계", "최종위험점수"]:
        if col in panel.columns:
            panel[col] = pd.to_numeric(panel[col], errors="coerce")

if alert is not None:
    for col in ["최종위험점수", "fta_ratio", "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]:
        if col in alert.columns:
            alert[col] = pd.to_numeric(alert[col], errors="coerce")

chains = sorted(panel["체인구분"].dropna().unique().tolist())
month_labels = (
    panel[["연월_sort", "연월_표시"]]
    .dropna()
    .drop_duplicates()
    .sort_values("연월_sort")
)
month_options = month_labels["연월_표시"].tolist()

# =========================================================
# 사이드바 메뉴
# =========================================================
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
    st.markdown("**결론 포인트:** 현재 어떤 체인이 얼마나 위험한지, 그리고 어떤 요인이 위험을 끌어올렸는지를 한눈에 보여줍니다.")

    c1, c2 = st.columns(2)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", chains)
    with c2:
        selected_month_label = st.selectbox("연월 선택", month_options, index=get_latest_index(month_options))

    panel_row = panel[
        (panel["체인구분"] == selected_chain) &
        (panel["연월_표시"] == selected_month_label)
    ]
    alert_row = alert[
        (alert["체인구분"] == selected_chain) &
        (alert["연월_표시"] == selected_month_label)
    ]

    if panel_row.empty or alert_row.empty:
        st.warning("선택한 체인/연월 데이터가 없어.")
        st.stop()

    p = panel_row.iloc[0]
    a = alert_row.iloc[0]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("최종위험점수", fmt_num(a.get("최종위험점수")))
    m2.metric("최종경보등급", f"{traffic_emoji(a.get('최종경보등급'))} {a.get('최종경보등급', '-')}")
    m3.metric("대체조달가능성", str(a.get("대체조달가능성", "-")))
    m4.metric("우선관리대상", str(a.get("상대적_우선관리대상", "-")))

    m5, m6, m7, m8 = st.columns(4)
    m5.metric("FTA 활용비중", fmt_pct(a.get("fta_ratio"), already_percent=True))
    m6.metric("상위1국의존도", fmt_pct(p.get("상위1국의존도"), already_percent=True))
    m7.metric("상위3국집중도", fmt_pct(p.get("상위3국집중도"), already_percent=True))
    m8.metric("HHI", fmt_num(p.get("HHI")))

    risk_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    if all(col in alert.columns for col in risk_cols):
        risk_df = pd.DataFrame({
            "리스크유형": ["가격", "수급", "물류", "정책이벤트"],
            "점수": [
                a.get("가격리스크점수"),
                a.get("수급리스크점수"),
                a.get("물류리스크점수"),
                a.get("정책이벤트리스크점수"),
            ]
        })

        st.markdown("### 리스크 유형별 점수 분해")
        fig = px.bar(
            risk_df,
            x="리스크유형",
            y="점수",
            color="리스크유형",
            text_auto=".2f",
            title=f"{selected_chain} / {selected_month_label} 리스크 유형별 기여도"
        )
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

        top_risk = risk_df.sort_values("점수", ascending=False).iloc[0]["리스크유형"]
    else:
        top_risk = "-"

    st.markdown("### 월별 최종위험점수 추이")
    chain_ts = alert[alert["체인구분"] == selected_chain].copy()
    chain_ts = chain_ts.sort_values("연월_sort")
    fig2 = px.line(
        chain_ts,
        x="연월_표시",
        y="최종위험점수",
        markers=True,
        title=f"{selected_chain} 월별 최종위험점수 추이"
    )
    fig2.update_layout(height=420, xaxis_title="연월", yaxis_title="최종위험점수")
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 발표용 자동 해석")
    st.write(
        f"- **{selected_month_label} 기준 {selected_chain}의 최종위험점수는 {fmt_num(a.get('최종위험점수'))}점이며, "
        f"경보등급은 `{a.get('최종경보등급', '-')}` 단계**로 나타났어."
    )
    st.write(
        f"- 상위 1개국 의존도는 **{fmt_pct(p.get('상위1국의존도'), already_percent=True)}**, "
        f"상위 3개국 집중도는 **{fmt_pct(p.get('상위3국집중도'), already_percent=True)}**, "
        f"HHI는 **{fmt_num(p.get('HHI'))}**로 확인돼 공급집중 구조를 함께 해석할 수 있어."
    )
    st.write(
        f"- 대체조달가능성은 **{a.get('대체조달가능성', '-')}**, "
        f"FTA 활용비중은 **{fmt_pct(a.get('fta_ratio'), already_percent=True)}** 수준이야."
    )
    st.write(f"- 해당 시점에서 상대적으로 가장 크게 작용한 리스크는 **{top_risk} 리스크**로 해석돼.")
    st.write(f"- 보정사유: **{a.get('보정사유', '-')}**")
    st.write(f"- 비고: **{a.get('비고', '-')}**")

# =========================================================
# 2. 품목·국가 취약성 분석
# =========================================================
elif menu == "2. 품목·국가 취약성 분석":
    st.header("🌍 품목별·국가별 취약성 진단")
    st.markdown("**결론 포인트:** 어느 품목이 어느 국가에 얼마나 집중되어 있고, 어떤 보정 요인 때문에 취약성이 높아졌는지 설명합니다.")

    c1, c2, c3 = st.columns(3)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", sorted(country["체인구분"].dropna().unique().tolist()))

    chain_months = (
        country[country["체인구분"] == selected_chain][["연월_sort", "연월_표시"]]
        .dropna()
        .drop_duplicates()
        .sort_values("연월_sort")
    )
    month_opts = chain_months["연월_표시"].tolist()

    with c2:
        selected_month_label = st.selectbox("연월 선택", month_opts, index=get_latest_index(month_opts))

    hs_candidates = (
        country[
            (country["체인구분"] == selected_chain) &
            (country["연월_표시"] == selected_month_label)
        ]["HS코드"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .dropna()
        .unique()
        .tolist()
    )
    hs_candidates = sorted(hs_candidates)

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
        st.warning("선택한 조건의 데이터가 없어.")
        st.stop()

    detail_df["국가별수입비중"] = pd.to_numeric(detail_df["국가별수입비중"], errors="coerce")
    detail_df = detail_df.sort_values("국가별수입비중", ascending=False)

    if not summary_df.empty:
        h = summary_df.iloc[0]
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("품목명", str(h.get("품목명", "-")))
        s2.metric("총수입금액", fmt_num(h.get("총수입금액"), 0))
        s3.metric("상위공급국", str(h.get("상위공급국", "-")))
        s4.metric("상위공급국비중", fmt_pct(h.get("상위공급국비중"), already_percent=False))

        s5, s6, s7, s8 = st.columns(4)
        s5.metric("수입국수", fmt_num(h.get("수입국수"), 0))
        s6.metric("평균수입단가", fmt_num(h.get("평균수입단가")))
        s7.metric("전월대비 수입금액증감률", fmt_pct(h.get("전월대비수입금액증감률"), already_percent=False))
        s8.metric("전년동월대비 수입금액증감률", fmt_pct(h.get("전년동월대비수입금액증감률"), already_percent=False))

    st.markdown("### 국가별 취약성 상세")
    show_cols = safe_columns(detail_df, [
        "국가명", "지역권", "FTA여부", "국가별수입비중", "지역권별수입비중",
        "상위공급국여부", "기본평가점수", "총보정점수", "최종보정점수", "최종판정", "비고"
    ])
    st.dataframe(detail_df[show_cols], use_container_width=True, height=420)

    if "국가명" in detail_df.columns and "국가별수입비중" in detail_df.columns:
        fig = px.bar(
            detail_df.head(10),
            x="국가명",
            y="국가별수입비중",
            color="지역권" if "지역권" in detail_df.columns else None,
            text_auto=".2f",
            title="상위 국가별 수입비중(%)"
        )
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    top_row = detail_df.iloc[0]
    st.markdown("### 발표용 자동 해석")
    st.write(
        f"- **{selected_month_label} 기준 HS {selected_hs} 품목은 `{top_row.get('국가명', '-')}`에 가장 크게 의존**하고 있으며, "
        f"해당 국가 수입비중은 **{fmt_pct(top_row.get('국가별수입비중'), already_percent=True)}**야."
    )
    if "FTA여부" in detail_df.columns:
        fta_n = detail_df[detail_df["FTA여부"] == "N"]
        st.write(f"- FTA 미체결 또는 비활용 국가 건수는 **{len(fta_n)}건**이야.")
    if "최종판정" in detail_df.columns:
        high_risk = detail_df[detail_df["최종판정"].isin(["높음", "매우높음"])]
        st.write(f"- 최종판정 기준 고위험 국가 수는 **{len(high_risk)}건**으로 확인돼.")
    if item_info is not None and "HS코드" in item_info.columns:
        info_df = item_info[item_info["HS코드"].astype(str).str.replace(".0", "", regex=False) == selected_hs]
        if not info_df.empty and "선정이유" in info_df.columns:
            st.write(f"- 품목 선정이유: **{info_df.iloc[0]['선정이유']}**")

# =========================================================
# 3. 충격 원인 추적
# =========================================================
elif menu == "3. 충격 원인 추적":
    st.header("📈 충격 원인 추적 및 선제대응 시사점")
    st.markdown("**결론 포인트:** 왜 그 달에 위험이 상승했는지 원인을 추적하고, 향후 몇 개월 내 어떤 대응을 검토해야 하는지 제시합니다.")

    c1, c2 = st.columns(2)
    with c1:
        selected_chain = st.selectbox("체인구분 선택", chains)
    with c2:
        selected_month_label = st.selectbox("기준 연월 선택", month_options, index=get_latest_index(month_options))

    panel_row = panel[
        (panel["체인구분"] == selected_chain) &
        (panel["연월_표시"] == selected_month_label)
    ]
    alert_row = alert[
        (alert["체인구분"] == selected_chain) &
        (alert["연월_표시"] == selected_month_label)
    ]

    if panel_row.empty or alert_row.empty:
        st.warning("선택한 조건의 데이터가 없어.")
        st.stop()

    a = alert_row.iloc[0]

    st.markdown("### 1) 환율 및 원자재 가격 추이")
    if market is not None:
        market_plot = market.sort_values("연월_sort").copy()
        value_cols = [c for c in ["환율", "월평균납가격", "월평균리튬가격", "월평균니켈가격"] if c in market_plot.columns]
        if value_cols:
            fig1 = px.line(
                market_plot,
                x="연월_표시",
                y=value_cols,
                title="환율 및 핵심 원자재 가격 추이"
            )
            fig1.update_layout(height=430, xaxis_title="연월")
            st.plotly_chart(fig1, use_container_width=True)

        market_sel = market_plot[market_plot["연월_표시"] == selected_month_label]
        if not market_sel.empty:
            mr = market_sel.iloc[0]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("환율", fmt_num(mr.get("환율")))
            c2.metric("납가격 정규화", fmt_num(mr.get("납가격정규화")))
            c3.metric("리튬가격 정규화", fmt_num(mr.get("리튬가격정규화")))
            c4.metric("니켈가격 정규화", fmt_num(mr.get("니켈가격정규화")))

    st.markdown("### 2) 정책이벤트 추이")
    if tpu is not None:
        tpu_plot = tpu.sort_values("연월_sort").copy()
        if "원천_TPU_INDEX" in tpu_plot.columns:
            fig2 = px.line(
                tpu_plot,
                x="연월_표시",
                y="원천_TPU_INDEX",
                markers=True,
                title="TPU(통상정책 불확실성) 추이"
            )
            fig2.update_layout(height=400, xaxis_title="연월")
            st.plotly_chart(fig2, use_container_width=True)

        tpu_sel = tpu_plot[tpu_plot["연월_표시"] == selected_month_label]
        if not tpu_sel.empty:
            tr = tpu_sel.iloc[0]
            st.write(f"- TPU 지수: **{fmt_num(tr.get('원천_TPU_INDEX'))}**")
            st.write(f"- 이벤트 보정: **{fmt_num(tr.get('이벤트보정'))}**")
            st.write(f"- 서사 배경: **{tr.get('서사배경', '-')}**")

    st.markdown("### 3) 물류 충격 추이")
    if gscpi is not None:
        gscpi_plot = gscpi.copy()
        gscpi_plot = gscpi_plot.rename(columns={" GSCPI": "GSCPI", " GSCPI_NORM": "GSCPI_NORM"})
        gscpi_plot = gscpi_plot.sort_values("연월_sort")

        if "GSCPI" in gscpi_plot.columns:
            fig3 = px.line(
                gscpi_plot,
                x="연월_표시",
                y="GSCPI",
                markers=True,
                title="GSCPI(글로벌 공급망 압력지수) 추이"
            )
            fig3.update_layout(height=400, xaxis_title="연월")
            st.plotly_chart(fig3, use_container_width=True)

        gs = gscpi_plot[gscpi_plot["연월_표시"] == selected_month_label]
        if not gs.empty:
            gr = gs.iloc[0]
            st.write(f"- GSCPI: **{fmt_num(gr.get('GSCPI'))}**")
            st.write(f"- GSCPI 정규화: **{fmt_num(gr.get('GSCPI_NORM'))}**")

    st.markdown("### 발표용 자동 해석")
    risk_df = pd.DataFrame({
        "유형": ["가격", "수급", "물류", "정책이벤트"],
        "점수": [
            a.get("가격리스크점수"),
            a.get("수급리스크점수"),
            a.get("물류리스크점수"),
            a.get("정책이벤트리스크점수"),
        ]
    }).sort_values("점수", ascending=False)

    dominant = risk_df.iloc[0]["유형"]
    st.write(
        f"- **{selected_month_label} 기준 {selected_chain}의 최종위험점수는 {fmt_num(a.get('최종위험점수'))}점**이며, "
        f"핵심 상승 요인은 **{dominant} 리스크**로 해석돼."
    )

    if dominant == "정책이벤트":
        st.write("- 시사점: 통상정책 변화와 지정학 이벤트가 선행 신호로 작용했을 가능성이 크며, **향후 2~3개월 내 조달선 재편과 대체국 확보 검토**가 필요해.")
    elif dominant == "물류":
        st.write("- 시사점: 글로벌 병목과 운송 차질이 압력 요인으로 작용한 만큼, **향후 1~2개월 내 선복·운송·통관 모니터링 강화**가 필요해.")
    elif dominant == "가격":
        st.write("- 시사점: 환율 또는 원자재 가격 변동이 수입단가에 직접 영향을 주고 있어, **선매입·계약단가 점검·가격헤지 검토**가 유효해.")
    else:
        st.write("- 시사점: 공급집중과 국가보정 요인이 결합된 구조이므로, **수입선 다변화와 FTA 활용 확대 전략**이 필요해.")

# =========================================================
# 4. 국가위험 근거 보기
# =========================================================
elif menu == "4. 국가위험 근거 보기":
    st.header("🛡️ 국가위험 점수의 설명 가능성")
    st.markdown("**결론 포인트:** 국가위험 점수는 임의의 숫자가 아니라, 분쟁·거버넌스·평화지표를 결합한 구조적 지표임을 보여줍니다.")

    if risk_master is None:
        st.warning("RISK_MASTER 시트를 찾지 못했어.")
        st.stop()

    country_options = sorted(risk_master["Country_std"].dropna().unique().tolist())
    year_options = sorted(risk_master["Year"].dropna().unique().tolist())

    c1, c2 = st.columns(2)
    with c1:
        selected_country = st.selectbox("국가 선택", country_options)
    with c2:
        selected_year = st.selectbox("연도 선택", year_options, index=get_latest_index(year_options))

    rr = risk_master[
        (risk_master["Country_std"] == selected_country) &
        (risk_master["Year"] == selected_year)
    ]

    if rr.empty:
        st.warning("선택한 국가/연도 데이터가 없어.")
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

    st.markdown("### 발표용 자동 해석")
    st.write(f"- **{selected_country} / {selected_year}년 종합위험점수는 {fmt_num(r.get('Composite_Risk_Score'))}점**이야.")
    st.write(f"- 최종 위험 판정은 **{r.get('Risk_Flag', '-')}** 이며, 사용 방식은 **{r.get('Composite_Method', '-')}** 이야.")
    if pd.notna(r.get("ACLED_Risk_Score")):
        st.write("- 최근 연도는 ACLED 기반 분쟁위험을 활용해 **현시점 위험 반영력**을 높였어.")
    else:
        st.write("- 과거 연도는 UCDP 기반 분쟁위험을 활용해 **연도별 비교 가능성**을 유지했어.")
    st.write("- 따라서 이 점수는 단순 직관이 아니라, **설명 가능한 구조적 위험지표**로 제시할 수 있어.")

# =========================================================
# 5. 기업 대응 시뮬레이터
# =========================================================
elif menu == "5. 기업 대응 시뮬레이터":
    st.header("🏭 기업 대응 전략 시뮬레이터")
    st.markdown("**결론 포인트:** 기업이 수입선 다변화, FTA 활용, 외부충격 대응 전략을 실제로 시험해볼 수 있는 실행형 도구입니다.")

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

    st.markdown("### 발표용 자동 권고")
    if diversify >= 10:
        st.write("- 수입선 다변화 폭이 의미 있는 수준이므로 **집중 리스크 완화 효과**를 기대할 수 있어.")
    else:
        st.write("- 다변화 폭이 작아 **구조적 위험 완화 효과는 제한적**일 수 있어.")
    if fta_change == "FTA 신규 활용":
        st.write("- FTA 신규 활용은 **통상비용 완충과 국가보정 리스크 완화**에 기여할 수 있어.")
    if shock_type == "정책이벤트 충격":
        st.write("- 정책이벤트 충격 가정하에서는 **2~3개월 내 조달선 재편 및 재고전략 점검**이 필요해.")
    elif shock_type == "물류충격":
        st.write("- 물류충격 가정하에서는 **1~2개월 내 선복 확보·통관 모니터링 강화**가 중요해.")
    elif shock_type == "가격충격":
        st.write("- 가격충격 가정하에서는 **단가 재협상·선매입·헤지 검토**가 중요해.")
    if new_top1 < 50:
        st.write("- 상위1국의존도를 50% 미만으로 낮추면 **구조적 취약성 완화**를 발표에서 강하게 제시할 수 있어.")

# =========================================================
# 6. 방법론 / 가중치 설명
# =========================================================
elif menu == "6. 방법론 / 가중치 설명":
    st.header("📘 방법론·가중치·재현 가능성 설명")
    st.markdown("**결론 포인트:** 본 모형은 어디에 고정가중치를 썼고, 어디에 규칙기반 보정을 썼고, 어디에 엔트로피를 썼는지 명확하게 설명할 수 있습니다.")

    st.markdown("""
### 1) 국가위험 기초점수
- `RISK_MASTER` 단계에서 **고정가중치 방식**을 사용
- 2021~2024년: `0.6 × UCDP_Risk_Score + 0.25 × WGI_Risk_Score_Used + 0.15 × GPI_Risk_Score`
- 2025년: `0.6 × ACLED_Risk_Score + 0.25 × WGI_Risk_Score_Used + 0.15 × GPI_Risk_Score`

### 2) 국가별 보정단계
- `COUNTRY_MONTHLY` 단계에서 **규칙기반 보정식**을 사용
- 공급국 집중, 지역권 집중, HHI, 상위공급국 여부, FTA 여부를 함께 반영

### 3) 최종 통합단계
- `PANEL_MONTHLY`, `ALERT_RESULT` 단계에서 **엔트로피 가중치**를 활용
- 가격 / 수급 / 물류 / 정책이벤트 리스크를 통합하여 최종위험점수를 도출

### 4) 왜 물류 / 정책이벤트가 100%로 보이는가?
- 각 카테고리 내부 변수가 1개뿐인 경우, **카테고리 내부 비중이 100%**로 표시됨
- 이는 전체 모델이 단일 변수라는 뜻이 아니라, **해당 카테고리 내부 구성상 유일한 변수**라는 의미야
    """)

    st.markdown("### ENTROPY_WEIGHT 시트")
    st.dataframe(entropy, use_container_width=True, height=360)

    st.markdown("### 체인별 비교표")
    st.dataframe(compare, use_container_width=True, height=240)

    st.markdown("### 발표용 한 줄 정리")
    st.write("- 본 모형은 **설명 가능성, 재현 가능성, 정책 활용 가능성**을 동시에 확보한 구조라고 정리할 수 있어.")
