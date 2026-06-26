import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="망보는사람들 공급망 리스크 대시보드",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# 공통 유틸
# =========================================================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def pick_sheet(xls, candidates):
    for s in candidates:
        if s in xls.sheet_names:
            return s
    return None

def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce")

def safe_ym(x, year=None, month=None):
    import pandas as pd
    import numpy as np

    if year is not None and month is not None and pd.notna(year) and pd.notna(month):
        try:
            return f"{int(year):04d}-{int(month):02d}"
        except:
            pass

    if pd.isna(x):
        return None

    if isinstance(x, pd.Timestamp):
        return f"{x.year:04d}-{x.month:02d}"

    if isinstance(x, str):
        s = x.strip()
        for sep in ["-", "/", "."]:
            parts = s.split(sep)
            if len(parts) >= 2 and parts[0].isdigit():
                try:
                    y = int(parts[0])
                    m = int(parts[1])
                    if 1 <= m <= 12:
                        return f"{y:04d}-{m:02d}"
                except:
                    pass
        try:
            dt = pd.to_datetime(s)
            return f"{dt.year:04d}-{dt.month:02d}"
        except:
            return s

    if isinstance(x, (int, np.integer)):
        s = str(int(x))
        if len(s) == 6 and s.isdigit():
            y = int(s[:4])
            m = int(s[4:])
            if 1 <= m <= 12:
                return f"{y:04d}-{m:02d}"
        return str(x)

    if isinstance(x, (float, np.floating)):
        y = int(x)
        frac = round(float(x) - y, 10)

        if abs(frac) < 1e-12:
            return f"{y:04d}-01"

        frac_str = f"{frac:.10f}".split(".")[1].rstrip("0")

        try:
            if len(frac_str) == 1:
                m = int(frac_str) * 10
            else:
                m = int(frac_str[:2])
            if 1 <= m <= 12:
                return f"{y:04d}-{m:02d}"
        except:
            pass

        m = int(round((float(x) - y) * 100))
        if 1 <= m <= 12:
            return f"{y:04d}-{m:02d}"

    return str(x)

def ensure_ym_column(df):
    df = df.copy()
    cols = list(df.columns)

    if "연월" in cols:
        df["연월_키"] = df["연월"].apply(safe_ym)
        df["연월"] = df["연월_키"]
        return df

    if "연도" in cols and "월" in cols:
        df["연월_키"] = df.apply(lambda r: safe_ym(None, r["연도"], r["월"]), axis=1)
        df["연월"] = df["연월_키"]
        return df

    if "연" in cols and "월" in cols:
        df["연월_키"] = df.apply(lambda r: safe_ym(None, r["연"], r["월"]), axis=1)
        df["연월"] = df["연월_키"]
        return df

    return df

def normalize_text_col(df, col):
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def find_first_existing(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def make_download_excel(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in data_dict.items():
            if df is not None and isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=sname[:31], index=False)
    return output.getvalue()

def safe_rank_pct(series):
    s = safe_numeric(series)
    if s.notna().sum() == 0:
        return pd.Series(np.zeros(len(series)), index=series.index)
    return s.rank(pct=True)

# =========================================================
# 데이터 로드
# =========================================================
@st.cache_data(show_spinner=False)
def load_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)

    sheet_map = {
        "country": pick_sheet(xls, ["NEW_COUNTRY_MONTHLY", "COUNTRY_MONTHLY"]),
        "panel": pick_sheet(xls, ["NEW_PANEL_MONTHLY", "PANEL_MONTHLY"]),
        "alert": pick_sheet(xls, ["NEW_ALERT_RESULT", "ALERT_RESULT"]),
        "chain_compare": pick_sheet(xls, ["NEW_체인별비교표", "체인별 비교표"]),
        "risk_master_clean": pick_sheet(xls, ["NEW_RISK_MASTER_CLEAN", "RISK_MASTER", "NEW_RISK_MASTER"]),
        "risk_fallback": pick_sheet(xls, ["NEW_RISK_FALLBACK", "RISK_FALLBACK"]),
        "entropy": pick_sheet(xls, ["NEW_ENTROPY_WEIGHT", "ENTROPY_WEIGHT"]),
        "norm_check": pick_sheet(xls, ["NEW_NORMALIZATION_CHECK", "NOMALIZATION_CHECK", "NORMALIZATION_CHECK"]),
        "method": pick_sheet(xls, ["NEW_METHOD_GUIDE", "METHOD_GUIDE"]),
        "market": pick_sheet(xls, ["MARKET_INDEX"]),
        "gscpi": pick_sheet(xls, ["GSCPI_INDEX"]),
        "tpu": pick_sheet(xls, ["TPU_INDEX"]),
        "hs_summary": pick_sheet(xls, ["HS_MONTHLY_SUMMARY"]),
        "data_info": pick_sheet(xls, ["DATA_INFO"]),
        "item_info": pick_sheet(xls, ["ITEM_INFO"]),
        "leadacid_raw": pick_sheet(xls, ["DATA_850710_납산축전지군", "DATA_850710_납산배터리", "DATA_850710_납산배터리군"]),
        "lithium_raw": pick_sheet(xls, ["DATA_850760_리튬이온배터리군"]),
    }

    data = {}
    for key, sname in sheet_map.items():
        if sname:
            df = pd.read_excel(uploaded_file, sheet_name=sname)
            df = clean_columns(df)
            df = ensure_ym_column(df)
            for txt_col in [
                "체인구분", "국가코드", "국가명", "FTA여부", "상위공급국여부",
                "최종경보등급", "최종판정", "대체조달가능성", "상대적_우선관리대상",
                "보정사유", "지역권"
            ]:
                df = normalize_text_col(df, txt_col)
            data[key] = df
        else:
            data[key] = None

    return xls.sheet_names, sheet_map, data

# =========================================================
# 검증 로직
# =========================================================
def run_checks(sheet_names, sheet_map, data):
    results = []

    def add(level, item, detail):
        results.append({"레벨": level, "점검항목": item, "상세": detail})

    required_keys = ["country", "panel", "alert"]
    for k in required_keys:
        if data[k] is None:
            add("FAIL", f"필수 시트 누락: {k}", f"{k} 시트를 찾지 못했습니다.")
        else:
            add("PASS", f"필수 시트 존재: {k}", f"{sheet_map[k]} 사용")

    for k in ["country", "panel", "alert", "market", "gscpi", "tpu", "hs_summary"]:
        df = data.get(k)
        if df is not None and "연월_키" in df.columns:
            miss = df["연월_키"].isna().sum()
            if miss == 0:
                add("PASS", f"{k} 연월 파싱", "모든 행 파싱 성공")
            else:
                add("WARN", f"{k} 연월 파싱", f"연월_키 결측 {miss}건")

    panel = data.get("panel")
    if panel is not None and "연월_키" in panel.columns:
        oct_cnt = panel["연월_키"].astype(str).str.endswith("-10").sum()
        if oct_cnt > 0:
            add("PASS", "10월 인식", f"-10 형식 {oct_cnt}건 확인")
        else:
            add("WARN", "10월 인식", "10월(-10) 데이터가 확인되지 않았습니다.")

    for key, cols in {
        "panel": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수"],
        "alert": ["대체조달가능성_점수", "최종위험점수"],
        "country": ["기본평가점수", "총보정점수", "최종보정점수"],
    }.items():
        df = data.get(key)
        if df is None:
            continue
        for c in cols:
            if c in df.columns:
                s = safe_numeric(df[c])
                bad = ((s < 0) | (s > 100)).sum()
                if bad == 0:
                    add("PASS", f"{key}.{c} 범위", "0~100 범위 정상")
                else:
                    add("WARN", f"{key}.{c} 범위", f"0~100 범위 이탈 {bad}건")

    if panel is not None and "fta_ratio" in panel.columns:
        s = safe_numeric(panel["fta_ratio"])
        bad = ((s < 0) | (s > 100)).sum()
        if bad == 0:
            add("PASS", "fta_ratio 범위", "0~100 범위 정상")
        else:
            add("FAIL", "fta_ratio 범위", f"0~100 범위 이탈 {bad}건")

    country = data.get("country")
    if country is not None and "국가코드" in country.columns:
        oth = country[country["국가코드"] == "OTH"].copy()
        if len(oth) == 0:
            add("WARN", "OTH 점검", "OTH 행이 없습니다.")
        else:
            if "최종보정점수" in oth.columns:
                filled = oth["최종보정점수"].notna().sum()
                if filled == 0:
                    add("PASS", "OTH 점검", "OTH는 최종보정점수 미부여")
                else:
                    add("FAIL", "OTH 점검", f"OTH 최종보정점수 부여 {filled}건")

    alert = data.get("alert")
    if panel is not None and alert is not None and "연월" in panel.columns and "체인구분" in panel.columns and "연월" in alert.columns and "체인구분" in alert.columns:
        pkey = set(panel["연월"].astype(str) + "||" + panel["체인구분"].astype(str))
        akey = set(alert["연월"].astype(str) + "||" + alert["체인구분"].astype(str))
        only_p = len(pkey - akey)
        only_a = len(akey - pkey)
        if only_p == 0 and only_a == 0:
            add("PASS", "panel-alert 키 연결", "연월+체인구분 기준 완전 일치")
        else:
            add("WARN", "panel-alert 키 연결", f"panel만 {only_p}건, alert만 {only_a}건")

    cc = data.get("chain_compare")
    if cc is not None and "종합_시사점" in cc.columns:
        add("PASS", "종합_시사점", "정량결과 기반 자동 생성 해석문구")

    return pd.DataFrame(results)

# =========================================================
# 앱 시작
# =========================================================
st.title("망보는사람들 공급망 리스크 대시보드")

uploaded_file = st.sidebar.file_uploader(
    "최종 엑셀 파일 업로드",
    type=["xlsx"],
    help="NEW 시트 포함 최종본 권장. 기존 시트만 있어도 자동 대응합니다."
)

menu = st.sidebar.radio(
    "메뉴",
    [
        "1. 종합 상황판",
        "2. 체인별 심층 분석",
        "3. 국가/공급선 상세 분석",
        "4. 충격 원인 추적",
        "5. 선행 신호 후보 탐지",
        "6. 기업 대응 시뮬레이터",
        "7. 대체국 추천 시스템",
        "8. 원천데이터 탐색 / 다운로드",
        "9. 데이터 검증 / 방법론",
    ]
)

show_note = st.sidebar.checkbox("종합 시사점 설명 표시", value=True)

if uploaded_file is None:
    st.info("왼쪽에서 최종 엑셀 파일(.xlsx)을 업로드해 주세요.")
    st.stop()

sheet_names, sheet_map, data = load_workbook(uploaded_file)
check_df = run_checks(sheet_names, sheet_map, data)

country = data["country"]
panel = data["panel"]
alert = data["alert"]
chain_compare = data["chain_compare"]
risk_master_clean = data["risk_master_clean"]
risk_fallback = data["risk_fallback"]
entropy = data["entropy"]
norm_check = data["norm_check"]
method = data["method"]
market = data["market"]
gscpi = data["gscpi"]
tpu = data["tpu"]
hs_summary = data["hs_summary"]
data_info = data["data_info"]
item_info = data["item_info"]
leadacid_raw = data["leadacid_raw"]
lithium_raw = data["lithium_raw"]

latest_month = None
if panel is not None and "연월" in panel.columns and panel["연월"].notna().any():
    latest_month = sorted(panel["연월"].dropna().astype(str).unique())[-1]

# =========================================================
# 1. 종합 상황판
# =========================================================
if menu == "1. 종합 상황판":
    st.subheader("종합 상황판")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("최신 기준월", latest_month if latest_month else "-")
    c2.metric("체인 수", int(panel["체인구분"].nunique()) if panel is not None and "체인구분" in panel.columns else 0)
    c3.metric("검증 FAIL", int((check_df["레벨"] == "FAIL").sum()))
    c4.metric("검증 WARN", int((check_df["레벨"] == "WARN").sum()))

    if panel is not None:
        latest_panel = panel.copy()
        if latest_month:
            latest_panel = latest_panel[latest_panel["연월"] == latest_month]

        st.markdown("#### 최신월 체인별 핵심 현황")
        cols_to_show = [c for c in [
            "체인구분", "최종위험점수", "최종경보등급",
            "가격리스크점수", "수급리스크점수", "물류리스크점수",
            "정책이벤트리스크점수", "fta_ratio", "지역권수"
        ] if c in latest_panel.columns]
        st.dataframe(latest_panel[cols_to_show], use_container_width=True)

        if "최종위험점수" in panel.columns:
            fig = px.line(
                panel.sort_values(["체인구분", "연월"]),
                x="연월", y="최종위험점수", color="체인구분",
                markers=True, title="체인별 최종위험점수 추이"
            )
            st.plotly_chart(fig, use_container_width=True)

        risk_cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in latest_panel.columns]
        if risk_cols:
            bar_df = latest_panel[["체인구분"] + risk_cols].melt(id_vars="체인구분", var_name="리스크축", value_name="점수")
            fig2 = px.bar(
                bar_df, x="리스크축", y="점수", color="체인구분",
                barmode="group", title=f"{latest_month} 4대 리스크 비교"
            )
            st.plotly_chart(fig2, use_container_width=True)

    if chain_compare is not None:
        st.markdown("#### 체인별 비교 요약")
        show_cols = [c for c in [
            "체인구분", "평균_최종위험점수", "평균_대체조달가능성점수",
            "우선관리대상_비중", "최고_위험_도달_시점", "종합_시사점"
        ] if c in chain_compare.columns]
        st.dataframe(chain_compare[show_cols], use_container_width=True)

        if show_note and "종합_시사점" in chain_compare.columns:
            st.caption("※ 종합_시사점은 별도 산식 점수가 아니라, 체인별 평균 4대 리스크 중 상대적으로 높은 축을 기준으로 자동 생성된 관리 해석문구입니다.")

# =========================================================
# 2. 체인별 심층 분석
# =========================================================
elif menu == "2. 체인별 심층 분석":
    st.subheader("체인별 심층 분석")

    if panel is None or "체인구분" not in panel.columns:
        st.warning("체인 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list)

    panel_c = panel[panel["체인구분"] == selected_chain].copy()
    alert_c = alert[alert["체인구분"] == selected_chain].copy() if alert is not None and "체인구분" in alert.columns else pd.DataFrame()
    country_c = country[country["체인구분"] == selected_chain].copy() if country is not None and "체인구분" in country.columns else pd.DataFrame()

    if not panel_c.empty:
        plot_cols = [c for c in ["최종위험점수", "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in panel_c.columns]
        fig = px.line(
            panel_c.sort_values("연월"),
            x="연월", y=plot_cols, markers=True,
            title=f"{selected_chain} 리스크 추이"
        )
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 최신월 핵심지표")
        if latest_month and not panel_c.empty:
            cur = panel_c[panel_c["연월"] == latest_month]
            cols = [c for c in [
                "연월", "체인구분", "최종위험점수", "최종경보등급",
                "경보점수기초", "국가보정합계", "fta_ratio", "지역권수"
            ] if c in cur.columns]
            st.dataframe(cur[cols], use_container_width=True)

    with c2:
        st.markdown("#### 최신월 상위 공급국")
        if latest_month and not country_c.empty:
            curc = country_c[country_c["연월"] == latest_month].copy()
            amount_col = find_first_existing(curc, ["국가수입금액", "수입금액", "금액", "수입액"])
            show_cols = [c for c in ["국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부", "기본평가점수", "최종보정점수"] if c in curc.columns]
            if amount_col:
                curc[amount_col] = safe_numeric(curc[amount_col])
                curc = curc.sort_values(amount_col, ascending=False)
                show_cols = [amount_col] + show_cols
            st.dataframe(curc[show_cols].head(10), use_container_width=True)

    if not alert_c.empty:
        st.markdown("#### 경보 결과")
        cols = [c for c in [
            "연월", "최종위험점수", "최종경보등급", "보정사유",
            "대체조달가능성_점수", "대체조달가능성", "상대적_우선관리대상", "비고"
        ] if c in alert_c.columns]
        st.dataframe(alert_c[cols].sort_values("연월"), use_container_width=True)

# =========================================================
# 3. 국가/공급선 상세 분석
# =========================================================
elif menu == "3. 국가/공급선 상세 분석":
    st.subheader("국가/공급선 상세 분석")

    if country is None:
        st.warning("COUNTRY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(country["체인구분"].dropna().unique().tolist()) if "체인구분" in country.columns else []
    selected_chain = st.selectbox("체인 선택", chain_list, key="country_chain")

    month_list = sorted(country["연월"].dropna().unique().tolist()) if "연월" in country.columns else []
    default_idx = len(month_list) - 1 if month_list else 0
    selected_month = st.selectbox("연월 선택", month_list, index=default_idx, key="country_month")

    sub = country[(country["체인구분"] == selected_chain) & (country["연월"] == selected_month)].copy()

    amount_col = find_first_existing(sub, ["국가수입금액", "수입금액", "금액", "수입액"])
    if amount_col:
        sub[amount_col] = safe_numeric(sub[amount_col])
        sub = sub.sort_values(amount_col, ascending=False)

    st.markdown("#### 국가별 상세 현황")
    cols = [c for c in [
        "국가코드", "국가명", "지역권", amount_col, "FTA여부", "상위공급국여부",
        "국가별수입비중", "지역권별수입비중", "집중도기여도",
        "기본평가점수", "총보정점수", "최종보정점수", "최종판정", "위험점수출처"
    ] if c and c in sub.columns]
    st.dataframe(sub[cols], use_container_width=True)

    if "최종보정점수" in sub.columns and "국가코드" in sub.columns:
        fig = px.bar(
            sub.head(15),
            x="국가코드",
            y="최종보정점수",
            color="최종보정점수",
            title=f"{selected_chain} / {selected_month} 국가별 최종보정점수"
        )
        st.plotly_chart(fig, use_container_width=True)

    if "FTA여부" in sub.columns and amount_col:
        fta_grp = sub.groupby("FTA여부")[amount_col].sum().reset_index()
        fig2 = px.pie(fta_grp, names="FTA여부", values=amount_col, title="FTA 여부별 수입금액 비중")
        st.plotly_chart(fig2, use_container_width=True)

# =========================================================
# 4. 충격 원인 추적
# =========================================================
elif menu == "4. 충격 원인 추적":
    st.subheader("충격 원인 추적")

    if panel is None:
        st.warning("PANEL 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="shock_chain")

    sub = panel[panel["체인구분"] == selected_chain].copy().sort_values("연월")
    if "최종위험점수" in sub.columns:
        sub["전월대비변화"] = safe_numeric(sub["최종위험점수"]).diff()

    cols = [c for c in ["연월", "최종위험점수", "전월대비변화", "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in sub.columns]
    st.dataframe(sub[cols], use_container_width=True)

    if "전월대비변화" in sub.columns:
        peak = sub.sort_values("전월대비변화", ascending=False).head(5)
        st.markdown("#### 급등 월 상위 5개")
        st.dataframe(peak[cols], use_container_width=True)

    latest = sub[sub["연월"] == latest_month] if latest_month else sub.tail(1)
    if not latest.empty:
        row = latest.iloc[0]
        risk_map = {}
        for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]:
            if c in latest.columns:
                risk_map[c] = row[c]
        if risk_map:
            top_cause = max(risk_map, key=lambda x: -999 if pd.isna(risk_map[x]) else risk_map[x])
            st.success(f"최신월({row['연월']}) 기준 상대적으로 가장 높은 충격 축: {top_cause}")

# =========================================================
# 5. 선행 신호 후보 탐지
# =========================================================
elif menu == "5. 선행 신호 후보 탐지":
    st.subheader("선행 신호 후보 탐지")

    if panel is None:
        st.warning("PANEL 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="lead_chain")

    sub = panel[panel["체인구분"] == selected_chain].copy().sort_values("연월")

    candidate_cols = [c for c in [
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
        "경보점수기초", "국가보정합계", "환율정규화", "납가격정규화", "리튬가격정규화",
        "니켈가격정규화", "GSCPI_Norm", "정책이벤트_raw", "최종위험점수"
    ] if c in sub.columns]

    if "최종위험점수" not in candidate_cols:
        st.warning("최종위험점수 컬럼이 없습니다.")
        st.stop()

    lead_rows = []
    y = safe_numeric(sub["최종위험점수"]).shift(-1)

    for c in candidate_cols:
        if c == "최종위험점수":
            continue
        x = safe_numeric(sub[c])
        corr0 = x.corr(safe_numeric(sub["최종위험점수"]))
        corr1 = x.corr(y)
        lead_rows.append([c, corr0, corr1])

    lead_df = pd.DataFrame(lead_rows, columns=["지표", "동월상관", "다음월최종위험상관"])
    lead_df["절대값_다음월"] = lead_df["다음월최종위험상관"].abs()
    lead_df = lead_df.sort_values("절대값_다음월", ascending=False)

    st.dataframe(lead_df, use_container_width=True)

    fig = px.bar(
        lead_df.head(10),
        x="지표", y="다음월최종위험상관",
        title=f"{selected_chain} 선행 신호 후보(다음월 최종위험과 상관)"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.caption("※ 탐색적 상관분석입니다. 인과관계 확정이 아니라 선행 후보 파악용입니다.")

# =========================================================
# 6. 기업 대응 시뮬레이터
# =========================================================
elif menu == "6. 기업 대응 시뮬레이터":
    st.subheader("기업 대응 시뮬레이터")

    if panel is None:
        st.warning("PANEL 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="sim_chain")

    base = panel[(panel["체인구분"] == selected_chain) & (panel["연월"] == latest_month)].copy()
    if base.empty:
        st.warning("최신월 데이터가 없습니다.")
        st.stop()

    row = base.iloc[0].copy()

    col1, col2 = st.columns(2)
    with col1:
        reduce_supply = st.slider("수급리스크 완화(%)", 0, 50, 10)
        reduce_price = st.slider("가격리스크 완화(%)", 0, 50, 10)
    with col2:
        reduce_logistics = st.slider("물류리스크 완화(%)", 0, 50, 10)
        reduce_policy = st.slider("정책리스크 완화(%)", 0, 50, 10)

    simulated = {}
    for c, r in {
        "가격리스크점수": reduce_price,
        "수급리스크점수": reduce_supply,
        "물류리스크점수": reduce_logistics,
        "정책이벤트리스크점수": reduce_policy,
    }.items():
        if c in base.columns:
            simulated[c] = row[c] * (1 - r / 100)

    weights = None
    if entropy is not None and {"단계", "변수명", "가중치"}.issubset(entropy.columns):
        wf = entropy[(entropy["단계"] == "PANEL_FINAL")].copy()
        if not wf.empty:
            weights = wf.set_index("변수명")["가중치"].to_dict()

    base_score = row["최종위험점수"] if "최종위험점수" in base.columns else np.nan

    if weights:
        sim_raw = 0
        used = 0
        for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]:
            if c in simulated and c in weights and pd.notna(simulated[c]):
                sim_raw += simulated[c] * weights[c]
                used += weights[c]
        sim_score = sim_raw / used if used > 0 else np.nan
    else:
        vals = [v for v in simulated.values() if pd.notna(v)]
        sim_score = np.mean(vals) if vals else np.nan

    c1, c2 = st.columns(2)
    c1.metric("현재 최종위험점수", "-" if pd.isna(base_score) else f"{base_score:.2f}")
    c2.metric("시뮬레이션 점수(참고)", "-" if pd.isna(sim_score) else f"{sim_score:.2f}")

    sim_df = pd.DataFrame({"구분": ["현재", "개선 후(참고)"], "점수": [base_score, sim_score]})
    fig = px.bar(sim_df, x="구분", y="점수", color="구분", title="대응 시뮬레이션 비교")
    st.plotly_chart(fig, use_container_width=True)

    st.caption("※ 최신월 4대 리스크 축 점수를 활용한 정책 실험용 참고 기능입니다.")

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
elif menu == "7. 대체국 추천 시스템":
    st.subheader("대체국 추천 시스템")

    if country is None:
        st.warning("COUNTRY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(country["체인구분"].dropna().unique().tolist()) if "체인구분" in country.columns else []
    if not chain_list:
        st.warning("체인구분 컬럼이 없습니다.")
        st.stop()

    selected_chain = st.selectbox("체인 선택", chain_list, key="alt_chain")
    cur = country[(country["체인구분"] == selected_chain) & (country["연월"] == latest_month)].copy()

    if cur.empty:
        st.warning("최신월 국가 데이터가 없습니다.")
        st.stop()

    cur["FTA가점"] = np.where(cur["FTA여부"] == "Y", 15, 0) if "FTA여부" in cur.columns else 0
    cur["비상위가점"] = np.where(cur["상위공급국여부"] == "N", 10, 0) if "상위공급국여부" in cur.columns else 0

    if "최종보정점수" in cur.columns:
        cur["안정성점수"] = 100 - safe_numeric(cur["최종보정점수"])
    elif "기본평가점수" in cur.columns:
        cur["안정성점수"] = 100 - safe_numeric(cur["기본평가점수"])
    else:
        cur["안정성점수"] = np.nan

    amount_col = find_first_existing(cur, ["국가수입금액", "수입금액", "금액", "수입액"])
    if amount_col:
        cur["규모가점"] = safe_rank_pct(cur[amount_col]) * 10
    else:
        cur["규모가점"] = 0

    cur["추천점수"] = cur["안정성점수"] * 0.7 + cur["FTA가점"] + cur["비상위가점"] + cur["규모가점"]
    if "국가코드" in cur.columns:
        cur = cur[cur["국가코드"] != "OTH"]
    cur = cur.sort_values("추천점수", ascending=False)

    show_cols = [c for c in [
        "국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부",
        "기본평가점수", "최종보정점수", "안정성점수", "추천점수"
    ] if c in cur.columns]
    st.dataframe(cur[show_cols].head(15), use_container_width=True)

    if "국가코드" in cur.columns:
        fig = px.bar(cur.head(10), x="국가코드", y="추천점수", color="추천점수", title=f"{selected_chain} 대체국 추천 상위 10")
        st.plotly_chart(fig, use_container_width=True)

    st.caption("※ 추천점수는 안정성, FTA 여부, 상위공급국 여부, 거래규모를 결합한 실무형 참고 점수입니다.")

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.subheader("원천데이터 탐색 / 다운로드")

    data_options = {
        "ITEM_INFO": item_info,
        "DATA_INFO": data_info,
        "HS_MONTHLY_SUMMARY": hs_summary,
        "COUNTRY_MONTHLY": country,
        "PANEL_MONTHLY": panel,
        "ALERT_RESULT": alert,
        "RISK_MASTER": risk_master_clean,
        "RISK_FALLBACK": risk_fallback,
        "MARKET_INDEX": market,
        "GSCPI_INDEX": gscpi,
        "TPU_INDEX": tpu,
        "납산배터리 raw": leadacid_raw,
        "리튬이온배터리 raw": lithium_raw,
    }

    available_names = [k for k, v in data_options.items() if v is not None]
    selected_name = st.selectbox("탐색할 데이터 선택", available_names)
    selected_df = data_options[selected_name]

    st.markdown(f"#### {selected_name} 미리보기")
    st.dataframe(selected_df.head(200), use_container_width=True)

    st.markdown("#### 기본 정보")
    c1, c2 = st.columns(2)
    c1.metric("행 수", len(selected_df))
    c2.metric("열 수", len(selected_df.columns))

    csv_bytes = selected_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label=f"{selected_name} CSV 다운로드",
        data=csv_bytes,
        file_name=f"{selected_name}.csv",
        mime="text/csv"
    )

# =========================================================
# 9. 데이터 검증 / 방법론
# =========================================================
elif menu == "9. 데이터 검증 / 방법론":
    st.subheader("데이터 검증 / 방법론")

    st.markdown("#### 1) 데이터 검증 결과")
    st.dataframe(check_df, use_container_width=True)

    fail_cnt = int((check_df["레벨"] == "FAIL").sum())
    warn_cnt = int((check_df["레벨"] == "WARN").sum())
    pass_cnt = int((check_df["레벨"] == "PASS").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("PASS", pass_cnt)
    c2.metric("WARN", warn_cnt)
    c3.metric("FAIL", fail_cnt)

    st.markdown("#### 2) 사용 시트 매핑")
    map_df = pd.DataFrame({"논리명": list(sheet_map.keys()), "실사용 시트": list(sheet_map.values())})
    st.dataframe(map_df, use_container_width=True)

    if entropy is not None:
        st.markdown("#### 3) 가중치 시트")
        st.dataframe(entropy, use_container_width=True)

    if norm_check is not None:
        st.markdown("#### 4) 정규화 점검 시트")
        st.dataframe(norm_check, use_container_width=True)

    if method is not None:
        st.markdown("#### 5) 방법론 시트")
        st.dataframe(method, use_container_width=True)

    st.markdown("#### 6) 검증 포함 엑셀 내보내기")
    export_dict = {k: v for k, v in data.items() if v is not None}
    export_dict["검증결과"] = check_df
    excel_bytes = make_download_excel(export_dict)

    st.download_button(
        "검증포함 엑셀 다운로드",
        data=excel_bytes,
        file_name="망보는사람들_dashboard_check.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.sidebar.markdown("---")
st.sidebar.caption("팀명: 망보는사람들")
st.sidebar.caption("최종 시트 기반 공급망 리스크 분석 앱")
