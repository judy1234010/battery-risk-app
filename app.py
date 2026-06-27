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

def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce")

def normalize_text_col(df, col):
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def find_first_existing(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def safe_rank_pct(series):
    s = safe_numeric(series)
    if s.notna().sum() == 0:
        return pd.Series(np.zeros(len(series)), index=series.index)
    return s.rank(pct=True)

def safe_ym(x, year=None, month=None):
    """
    10월 깨짐 방지용 연월 파서
    예:
    - 2021.01 -> 2021-01
    - 2021.1  -> 2021-10
    - 202110  -> 2021-10
    """
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
            if len(parts) >= 2 and str(parts[0]).isdigit():
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

        try:
            m = int(round((float(x) - y) * 100))
            if 1 <= m <= 12:
                return f"{y:04d}-{m:02d}"
        except:
            pass

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

def make_download_excel(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in data_dict.items():
            if df is not None and isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=sname[:31], index=False)
    return output.getvalue()

def classify_signal(corr):
    if pd.isna(corr):
        return "해석유보"
    a = abs(corr)
    if a >= 0.6 and corr > 0:
        return "강한 양(+) 선행 후보"
    elif a >= 0.6 and corr < 0:
        return "강한 음(-) 반대 신호"
    elif a >= 0.4 and corr > 0:
        return "중간 양(+) 후보"
    elif a >= 0.4 and corr < 0:
        return "중간 음(-) 후보"
    else:
        return "참고 수준"

def pick_existing_sheet(sheet_names, candidates):
    for s in candidates:
        if s in sheet_names:
            return s
    return None

def get_col_case_insensitive(df, candidates):
    lower_map = {str(c).lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return None

def score_band_3(v, high=70, mid=40):
    if pd.isna(v):
        return "해석유보"
    if v >= high:
        return "취약"
    elif v >= mid:
        return "주의"
    else:
        return "양호"

def inverse_score_band(v, low=30, mid=60):
    """
    낮을수록 취약한 지표용
    예: fta_ratio, 대체조달가능성_점수
    """
    if pd.isna(v):
        return "해석유보"
    if v <= low:
        return "취약"
    elif v <= mid:
        return "주의"
    else:
        return "양호"

def safe_fmt(v, digits=1):
    if pd.isna(v):
        return "-"
    return f"{float(v):.{digits}f}"

# =========================================================
# 데이터 로드
# =========================================================
@st.cache_data(show_spinner=False)
def load_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    sheet_map = {
        "item_info": pick_existing_sheet(sheet_names, ["ITEM_INFO"]),
        "data_info": pick_existing_sheet(sheet_names, ["DATA_INFO"]),
        "risk_master": pick_existing_sheet(sheet_names, ["RISK_MASTER"]),
        "risk_fallback": pick_existing_sheet(sheet_names, ["RISK_FALLBACK"]),
        "country": pick_existing_sheet(sheet_names, ["COUNTRY_MONTHLY"]),
        "market": pick_existing_sheet(sheet_names, ["MARKET_INDEX"]),
        "gscpi": pick_existing_sheet(sheet_names, ["GSCPI_INDEX"]),
        "tpu": pick_existing_sheet(sheet_names, ["TPU_INDEX"]),
        "hs_summary": pick_existing_sheet(sheet_names, ["HS_MONTHLY_SUMMARY"]),
        "panel": pick_existing_sheet(sheet_names, ["PANEL_MONTHLY"]),
        "alert": pick_existing_sheet(sheet_names, ["ALERT_RESULT"]),
        "chain_compare": pick_existing_sheet(sheet_names, ["체인별 비교표"]),
        "entropy": pick_existing_sheet(sheet_names, ["ENTROPY_WEIGHT"]),
        "norm_check": pick_existing_sheet(sheet_names, ["NOMALIZATION_CHECK", "NORMALIZATION_CHECK"]),
        "method": pick_existing_sheet(sheet_names, ["METHOD_GUIDE"]),
        "leadacid_raw": pick_existing_sheet(sheet_names, ["DATA_850710_납산배터리", "DATA_850710_납산배터리군"]),
        "lithium_raw": pick_existing_sheet(sheet_names, ["DATA_850760_리튬이온배터리군"]),
    }

    data = {}

    for key, sname in sheet_map.items():
        if sname is None:
            data[key] = None
            continue

        df = pd.read_excel(uploaded_file, sheet_name=sname)
        df = clean_columns(df)
        df = ensure_ym_column(df)

        for txt_col in [
            "체인구분", "국가코드", "국가명", "FTA여부", "상위공급국여부",
            "최종경보등급", "최종판정", "대체조달가능성", "상대적_우선관리대상",
            "보정사유", "지역권", "품목명", "품목군", "배터리유형"
        ]:
            df = normalize_text_col(df, txt_col)

        data[key] = df

    return sheet_names, sheet_map, data

# =========================================================
# 검증 로직
# =========================================================
def run_checks(sheet_names, sheet_map, data):
    results = []

    def add(level, item, detail):
        results.append({"레벨": level, "점검항목": item, "상세": detail})

    required = ["country", "panel", "alert", "chain_compare"]
    for k in required:
        if data.get(k) is None:
            add("FAIL", f"필수 시트 누락: {k}", f"{k} 시트를 찾지 못했습니다.")
        else:
            add("PASS", f"필수 시트 존재: {k}", f"{sheet_map[k]} 사용")

    gscpi = data.get("gscpi")
    if gscpi is not None:
        cols = list(gscpi.columns)
        if "GSCPI" in cols and "GSCPI_NORM" in cols:
            add("PASS", "GSCPI 컬럼명", "GSCPI / GSCPI_NORM 정상")
        else:
            add("WARN", "GSCPI 컬럼명", f"현재 컬럼: {cols}")

    for k in ["data_info", "country", "market", "gscpi", "tpu", "hs_summary", "panel", "alert", "leadacid_raw", "lithium_raw"]:
        df = data.get(k)
        if df is not None and "연월_키" in df.columns:
            miss = df["연월_키"].isna().sum()
            if miss == 0:
                add("PASS", f"{k} 연월 파싱", "모든 행 파싱 성공")
            else:
                add("WARN", f"{k} 연월 파싱", f"연월_키 결측 {miss}건")

    panel = data.get("panel")
    if panel is not None and "연월" in panel.columns:
        oct_cnt = panel["연월"].astype(str).str.endswith("-10").sum()
        if oct_cnt > 0:
            add("PASS", "10월 인식", f"-10 형식 {oct_cnt}건 확인")
        else:
            add("WARN", "10월 인식", "10월(-10) 데이터 미확인")

    score_cols_by_table = {
        "panel": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수"],
        "alert": ["대체조달가능성_점수", "최종위험점수"],
        "country": ["기본평가점수", "총보정점수", "최종보정점수"],
    }

    for table, cols in score_cols_by_table.items():
        df = data.get(table)
        if df is None:
            continue
        for c in cols:
            if c in df.columns:
                s = safe_numeric(df[c])
                bad = ((s < 0) | (s > 100)).sum()
                if bad == 0:
                    add("PASS", f"{table}.{c} 범위", "0~100 정상")
                else:
                    add("WARN", f"{table}.{c} 범위", f"0~100 이탈 {bad}건")

    if panel is not None and "fta_ratio" in panel.columns:
        s = safe_numeric(panel["fta_ratio"])
        bad = ((s < 0) | (s > 100)).sum()
        if bad == 0:
            add("PASS", "fta_ratio 범위", "0~100 정상")
        else:
            add("FAIL", "fta_ratio 범위", f"0~100 이탈 {bad}건")

    country = data.get("country")
    if country is not None and "국가코드" in country.columns:
        oth = country[country["국가코드"] == "OTH"].copy()
        if len(oth) == 0:
            add("WARN", "OTH 점검", "OTH 행이 없음")
        else:
            if "최종보정점수" in oth.columns:
                filled = oth["최종보정점수"].notna().sum()
                if filled == 0:
                    add("PASS", "OTH 점검", "OTH 최종보정점수 미부여")
                else:
                    add("FAIL", "OTH 점검", f"OTH 최종보정점수 입력 {filled}건")

    alert = data.get("alert")
    if panel is not None and alert is not None and "연월" in panel.columns and "체인구분" in panel.columns and "연월" in alert.columns and "체인구분" in alert.columns:
        pkey = set(panel["연월"].astype(str) + "||" + panel["체인구분"].astype(str))
        akey = set(alert["연월"].astype(str) + "||" + alert["체인구분"].astype(str))
        only_p = len(pkey - akey)
        only_a = len(akey - pkey)
        if only_p == 0 and only_a == 0:
            add("PASS", "panel-alert 연결", "연월+체인구분 완전 일치")
        else:
            add("WARN", "panel-alert 연결", f"panel만 {only_p}건 / alert만 {only_a}건")

    cc = data.get("chain_compare")
    if cc is not None and "우선관리대상_비중" in cc.columns:
        add("PASS", "체인별 비교표", "우선관리대상_비중 컬럼 확인")

    return pd.DataFrame(results)

# =========================================================
# 앱 시작
# =========================================================
st.title("망보는사람들 공급망 리스크 대시보드")

uploaded_file = st.sidebar.file_uploader(
    "최종 엑셀 파일 업로드",
    type=["xlsx"],
    help="망보는사람들 최종 시트 기반 파일을 업로드하세요."
)

menu = st.sidebar.radio(
    "메뉴",
    [
        "1. 종합 상황판",
        "2. 체인별 심층 분석",
        "3. 국가/공급선 상세 분석",
        "4. 충격 원인 추적",
        "5. 선행 신호 후보 탐지",
        "6. 기업 대응 우선순위 추천 / 시뮬레이터",
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

item_info = data["item_info"]
data_info = data["data_info"]
risk_master = data["risk_master"]
risk_fallback = data["risk_fallback"]
country = data["country"]
market = data["market"]
gscpi = data["gscpi"]
tpu = data["tpu"]
hs_summary = data["hs_summary"]
panel = data["panel"]
alert = data["alert"]
chain_compare = data["chain_compare"]
entropy = data["entropy"]
norm_check = data["norm_check"]
method = data["method"]
leadacid_raw = data["leadacid_raw"]
lithium_raw = data["lithium_raw"]

latest_month = None
min_month = None
max_month = None
month_list_global = []

if panel is not None and "연월" in panel.columns and panel["연월"].notna().any():
    month_list_global = sorted(panel["연월"].dropna().astype(str).unique().tolist())
    if month_list_global:
        latest_month = month_list_global[-1]
        min_month = month_list_global[0]
        max_month = month_list_global[-1]

# =========================================================
# 1. 종합 상황판
# =========================================================
if menu == "1. 종합 상황판":
    st.subheader("종합 상황판")
    st.info(f"""
이 메뉴는 전체 공급망 리스크 현황을 한눈에 보는 요약 화면입니다.

활용 방법
1) 선택 기준월 체인별 최종위험 수준을 확인합니다.
2) 4대 리스크축(가격·수급·물류·정책) 중 어떤 축이 상대적으로 높은지 비교합니다.
3) 체인별 비교 요약을 통해 평균 위험수준, 대체조달 여건, 우선관리 필요성을 함께 봅니다.

참고
- 본 화면의 기본 기준월은 **업로드된 데이터의 최신월({latest_month if latest_month else '-'})** 입니다.
- 즉, 시스템 현재 날짜가 아니라 **업로드 데이터 기준 최신 시점 분석**입니다.
- 데이터 범위: **{min_month if min_month else '-'} ~ {max_month if max_month else '-'}**
""")

    selected_month = latest_month
    if month_list_global:
        selected_month = st.selectbox(
            "기준월 선택",
            month_list_global,
            index=len(month_list_global) - 1,
            key="summary_month"
        )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("업로드 데이터 최신월", latest_month if latest_month else "-")
    c2.metric("현재 선택 기준월", selected_month if selected_month else "-")
    c3.metric("검증 WARN", int((check_df["레벨"] == "WARN").sum()))
    c4.metric("검증 FAIL", int((check_df["레벨"] == "FAIL").sum()))

    st.markdown("#### 4대 리스크축 산정 기준")
    risk_basis_df = pd.DataFrame({
        "리스크축": ["가격리스크", "수급리스크", "물류리스크", "정책이벤트리스크"],
        "주요 산정 지표": [
            "환율정규화, 납가격정규화, 리튬가격정규화, 니켈가격정규화",
            "상위1국의존도, HHI, 수입국수, CV, 국가보정합계",
            "GSCPI_Norm",
            "TPU_INDEX의 이벤트보정값(정책이벤트_raw)"
        ],
        "해석": [
            "환율·원자재 가격 변동에 따른 가격 충격 수준",
            "공급집중도와 조달 불안정성 수준",
            "글로벌 공급망 병목·운송 압력 수준",
            "통상·정책 이벤트 충격 수준"
        ]
    })
    st.dataframe(risk_basis_df, use_container_width=True)

    if panel is not None:
        selected_panel = panel.copy()
        if selected_month:
            selected_panel = selected_panel[selected_panel["연월"] == selected_month]

        st.markdown("#### 선택월 체인별 핵심 현황")
        cols_to_show = [c for c in [
            "체인구분", "최종위험점수", "최종경보등급",
            "가격리스크점수", "수급리스크점수", "물류리스크점수",
            "정책이벤트리스크점수", "fta_ratio", "지역권수"
        ] if c in selected_panel.columns]
        st.dataframe(selected_panel[cols_to_show], use_container_width=True)

        if "최종위험점수" in panel.columns:
            fig = px.line(
                panel.sort_values(["체인구분", "연월"]),
                x="연월",
                y="최종위험점수",
                color="체인구분",
                markers=True,
                title="체인별 최종위험점수 추이"
            )
            st.plotly_chart(fig, use_container_width=True)

        risk_cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in selected_panel.columns]
        if risk_cols and not selected_panel.empty:
            bar_df = selected_panel[["체인구분"] + risk_cols].melt(id_vars="체인구분", var_name="리스크축", value_name="점수")
            fig2 = px.bar(
                bar_df,
                x="리스크축",
                y="점수",
                color="체인구분",
                barmode="group",
                title=f"{selected_month} 4대 리스크 비교"
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
            st.caption("※ 종합_시사점은 별도 점수가 아니라, 4대 리스크 평균 중 상대적으로 높은 축을 기준으로 자동 생성된 관리 해석문구입니다.")

        if "체인구분" in chain_compare.columns and "우선관리대상_비중" in chain_compare.columns:
            lithium_rows = chain_compare[chain_compare["체인구분"].astype(str).str.contains("리튬", na=False)]
            if not lithium_rows.empty:
                val = lithium_rows["우선관리대상_비중"].iloc[0]
                try:
                    if pd.notna(val) and float(val) == 0:
                        st.caption("""
리튬이온배터리군의 우선관리대상 비중이 0%라는 것은 위험이 없다는 뜻이 아니라,
분석기간 내 '고위험'과 '저대체조달'이 동시에 강하게 충족된 월이 없었다는 의미입니다.
즉 위험 상승 구간은 있었지만, 우선관리 판정 규칙상 동시충족 빈도가 낮았던 것입니다.
""")
                except:
                    pass

# =========================================================
# 2. 체인별 심층 분석
# =========================================================
elif menu == "2. 체인별 심층 분석":
    st.subheader("체인별 심층 분석")
    st.info("""
이 메뉴는 특정 체인을 선택해 월별 리스크 추이와 공급구조를 자세히 보는 화면입니다.

활용 방법
1) 체인별 최종위험점수와 4대 리스크축의 시계열 변화를 봅니다.
2) 최신월 기준 상위 공급국과 국가보정 수준을 확인합니다.
3) 경보등급, 대체조달가능성, 우선관리대상 여부를 함께 해석합니다.
""")

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
            x="연월",
            y=plot_cols,
            markers=True,
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
    st.info("""
이 메뉴는 특정 체인과 특정 월을 기준으로 국가별 공급선 구조를 확인하는 화면입니다.

활용 방법
1) 어떤 국가가 많이 들어오는지, 수입비중과 지역권 구조를 확인합니다.
2) 국가별 기본위험과 구조보정이 반영된 최종보정점수를 비교합니다.
3) FTA 여부, 상위공급국 여부를 함께 보며 공급선 다변화 여지를 판단합니다.
""")

    if country is None:
        st.warning("COUNTRY_MONTHLY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(country["체인구분"].dropna().unique().tolist()) if "체인구분" in country.columns else []
    if not chain_list:
        st.warning("체인 목록을 찾을 수 없습니다.")
        st.stop()

    selected_chain = st.selectbox("체인 선택", chain_list, key="country_chain")

    month_list = sorted(country["연월"].dropna().unique().tolist()) if "연월" in country.columns else []
    if not month_list:
        st.warning("연월 정보가 없습니다.")
        st.stop()

    default_idx = len(month_list) - 1
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

    st.caption("※ 국가별 최종보정점수는 기본 국가위험에 실제 수입구조 보정을 반영한 국가단위 위험지표입니다.")

    if "최종보정점수" in sub.columns and "국가코드" in sub.columns:
        plot_df = sub.dropna(subset=["국가코드"]).copy()
        fig = px.bar(
            plot_df.head(15),
            x="국가코드",
            y="최종보정점수",
            color="최종보정점수",
            color_continuous_scale="Reds",
            title=f"{selected_chain} / {selected_month} 국가별 최종보정점수"
        )
        st.plotly_chart(fig, use_container_width=True)

# =========================================================
# 4. 충격 원인 추적
# =========================================================
elif menu == "4. 충격 원인 추적":
    st.subheader("충격 원인 추적")
    st.info("""
이 메뉴는 특정 체인에서 위험점수가 크게 움직인 시점과 그 원인을 추적하는 화면입니다.

활용 방법
1) 전월 대비 최종위험점수 상승폭이 큰 월을 우선 식별합니다.
2) 음(-)의 큰 변화는 위험 완화 구간을 의미하므로 별도로 해석할 수 있습니다.
3) 해당 시점에 가격·수급·물류·정책 중 어떤 축이 상대적으로 높았는지 확인합니다.
4) 정책이벤트리스크가 높았던 월은 TPU_INDEX 기반 정책 이벤트 배경 코멘트를 함께 참고합니다.
""")

    if panel is None:
        st.warning("PANEL_MONTHLY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="shock_chain")

    sub = panel[panel["체인구분"] == selected_chain].copy().sort_values("연월")

    if "최종위험점수" in sub.columns:
        sub["전월대비변화"] = safe_numeric(sub["최종위험점수"]).diff()

    cols = [c for c in [
        "연월", "최종위험점수", "전월대비변화",
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"
    ] if c in sub.columns]
    st.dataframe(sub[cols], use_container_width=True)

    selected_shock_month = None

    if "전월대비변화" in sub.columns:
        up_df = sub[sub["전월대비변화"].notna()].sort_values("전월대비변화", ascending=False).head(5)
        down_df = sub[sub["전월대비변화"].notna()].sort_values("전월대비변화", ascending=True).head(5)

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 급등 월 상위 5개")
            st.dataframe(up_df[cols], use_container_width=True)
        with c2:
            st.markdown("#### 급락 월 상위 5개")
            st.dataframe(down_df[cols], use_container_width=True)

        shock_month_candidates = sub["연월"].dropna().astype(str).tolist()
        default_shock_month = latest_month if latest_month in shock_month_candidates else shock_month_candidates[-1]
        selected_shock_month = st.selectbox(
            "정책 이벤트 배경을 확인할 연월 선택",
            shock_month_candidates,
            index=shock_month_candidates.index(default_shock_month),
            key="shock_month"
        )

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

    # TPU_INDEX 기반 정책이벤트 배경 코멘트
    if selected_shock_month and tpu is not None and "연월" in tpu.columns:
        tpu_sub = tpu[tpu["연월"].astype(str) == str(selected_shock_month)].copy()

        st.markdown("#### 정책 이벤트 배경 코멘트")
        if not tpu_sub.empty:
            tpu_row = tpu_sub.iloc[0]
            event_score = tpu_row["이벤트보정"] if "이벤트보정" in tpu_row.index else np.nan
            raw_tpu = tpu_row["원천_TPU_INDEX"] if "원천_TPU_INDEX" in tpu_row.index else np.nan
            narrative = tpu_row["서사배경"] if "서사배경" in tpu_row.index else ""

            st.info(f"""
- 기준 연월: **{selected_shock_month}**
- 원천 TPU_INDEX: **{safe_fmt(raw_tpu, 2)}**
- 이벤트보정값(정책이벤트리스크 원천): **{safe_fmt(event_score, 4)}**

**해석 참고**
{narrative if narrative else '해당 연월 서사배경 정보가 없습니다.'}

※ 위 코멘트는 TPU_INDEX 기반 정량지표를 해석하기 위한 보조설명입니다.
""")
        else:
            st.warning("선택한 연월에 해당하는 TPU_INDEX 배경 정보가 없습니다.")

# =========================================================
# 5. 선행 신호 후보 탐지
# =========================================================
elif menu == "5. 선행 신호 후보 탐지":
    st.subheader("선행 신호 후보 탐지")
    st.info("""
이 메뉴는 '다음 달 최종위험점수'와 먼저 같이 움직이는 지표를 찾는 탐색 화면입니다.

읽는 방법
1) '다음월최종위험상관'의 절댓값이 클수록 선행 신호 후보 가능성이 큽니다.
2) 양(+)이면 해당 지표 상승 뒤 다음 달 위험 상승 경향,
   음(-)이면 해당 지표 상승 뒤 다음 달 위험 하락 경향을 뜻합니다.
3) 본 결과는 인과관계 확정이 아니라, 선제 모니터링 우선지표를 찾기 위한 참고용입니다.
""")

    if panel is None:
        st.warning("PANEL_MONTHLY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="lead_chain")

    sub = panel[panel["체인구분"] == selected_chain].copy().sort_values("연월")

    gscpi_norm_col = get_col_case_insensitive(sub, ["GSCPI_Norm", "GSCPI_NORM"])
    policy_raw_col = get_col_case_insensitive(sub, ["정책이벤트_raw", "정책이벤트_RAW"])

    candidate_cols = [c for c in [
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
        "경보점수기초", "국가보정합계", "환율정규화", "납가격정규화",
        "리튬가격정규화", "니켈가격정규화", "최종위험점수"
    ] if c in sub.columns]

    if gscpi_norm_col and gscpi_norm_col not in candidate_cols:
        candidate_cols.append(gscpi_norm_col)
    if policy_raw_col and policy_raw_col not in candidate_cols:
        candidate_cols.append(policy_raw_col)

    if "최종위험점수" not in candidate_cols:
        st.warning("최종위험점수 컬럼이 없습니다.")
        st.stop()

    lead_rows = []
    y_next = safe_numeric(sub["최종위험점수"]).shift(-1)

    for c in candidate_cols:
        if c == "최종위험점수":
            continue
        x = safe_numeric(sub[c])
        corr_now = x.corr(safe_numeric(sub["최종위험점수"]))
        corr_next = x.corr(y_next)
        lead_rows.append([c, corr_now, corr_next])

    lead_df = pd.DataFrame(lead_rows, columns=["지표", "동월상관", "다음월최종위험상관"])
    lead_df["절대값_다음월"] = lead_df["다음월최종위험상관"].abs()
    lead_df["해석"] = lead_df["다음월최종위험상관"].apply(classify_signal)
    lead_df = lead_df.sort_values("절대값_다음월", ascending=False)

    st.dataframe(lead_df, use_container_width=True)

    fig = px.bar(
        lead_df.head(10),
        x="지표",
        y="다음월최종위험상관",
        color="다음월최종위험상관",
        color_continuous_scale="RdBu_r",
        title=f"{selected_chain} 선행 신호 후보(다음월 최종위험 상관)"
    )
    st.plotly_chart(fig, use_container_width=True)

    st.caption("※ 상관값이 크다고 해서 인과관계가 확정되는 것은 아니며, 선제 모니터링 후보를 찾기 위한 탐색적 기능입니다.")
    
# =========================================================
# 6. 기업 대응 우선순위 추천 / 시뮬레이터
# =========================================================
elif menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.subheader("기업 대응 우선순위 추천 / 시뮬레이터")
    st.info(f"""
이 메뉴는 선택 기준월 체인의 취약요인을 진단하고,
객관지표 기반으로 우선 대응전략을 추천한 뒤,
선택한 대응조치 적용 시 리스크 완화 효과를 시뮬레이션하는 기능입니다.

진행 순서
1) 현재 취약점 진단
2) 대응 우선순위 자동 추천
3) 대응 강도 조정 후 개선 효과 시뮬레이션

참고
- 기본 기준월은 업로드된 데이터의 최신월인 **{latest_month if latest_month else '-'}** 입니다.
- 즉, 시스템 현재 날짜가 아니라 **업로드 데이터 기준 최신 관측월 분석**입니다.
- 데이터 범위: **{min_month if min_month else '-'} ~ {max_month if max_month else '-'}**
""")

    if panel is None:
        st.warning("PANEL_MONTHLY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="sim_chain")

    selected_month = latest_month
    if month_list_global:
        selected_month = st.selectbox(
            "분석 기준월 선택",
            month_list_global,
            index=len(month_list_global) - 1,
            key="sim_month"
        )

    base = panel[(panel["체인구분"] == selected_chain) & (panel["연월"] == selected_month)].copy()
    if base.empty:
        st.warning("선택 기준월 데이터가 없습니다.")
        st.stop()

    row = base.iloc[0].copy()

    base_alert = pd.DataFrame()
    if alert is not None and {"체인구분", "연월"}.issubset(alert.columns):
        base_alert = alert[(alert["체인구분"] == selected_chain) & (alert["연월"] == selected_month)].copy()

    base_country = pd.DataFrame()
    if country is not None and {"체인구분", "연월"}.issubset(country.columns):
        base_country = country[(country["체인구분"] == selected_chain) & (country["연월"] == selected_month)].copy()

    # -----------------------------
    # 1) 취약점 진단
    # -----------------------------
    price_score = row["가격리스크점수"] if "가격리스크점수" in row.index else np.nan
    supply_score = row["수급리스크점수"] if "수급리스크점수" in row.index else np.nan
    logistics_score = row["물류리스크점수"] if "물류리스크점수" in row.index else np.nan
    policy_score = row["정책이벤트리스크점수"] if "정책이벤트리스크점수" in row.index else np.nan
    fta_ratio = row["fta_ratio"] if "fta_ratio" in row.index else np.nan
    region_cnt = row["지역권수"] if "지역권수" in row.index else np.nan

    alt_score = np.nan
    if not base_alert.empty and "대체조달가능성_점수" in base_alert.columns:
        alt_score = base_alert["대체조달가능성_점수"].iloc[0]

    top1_share = np.nan
    if not base_country.empty:
        share_col = find_first_existing(base_country, ["국가별수입비중", "수입비중"])
        if share_col:
            top1_share = safe_numeric(base_country[share_col]).max()

    hhi = np.nan
    country_cnt = np.nan
    if not base_country.empty:
        share_col = find_first_existing(base_country, ["국가별수입비중", "수입비중"])
        if share_col:
            shares = safe_numeric(base_country[share_col]).dropna() / 100
            if len(shares) > 0:
                hhi = (shares.pow(2).sum()) * 10000
        code_col = find_first_existing(base_country, ["국가코드"])
        if code_col:
            country_cnt = base_country[code_col].replace("OTH", np.nan).dropna().nunique()

    diag_df = pd.DataFrame({
        "항목": [
            "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
            "FTA 비중", "대체조달가능성_점수", "상위1국 의존도", "HHI", "수입국 수", "지역권 수"
        ],
        "현재값": [
            price_score, supply_score, logistics_score, policy_score,
            fta_ratio, alt_score, top1_share, hhi, country_cnt, region_cnt
        ],
        "판정": [
            score_band_3(price_score),
            score_band_3(supply_score),
            score_band_3(logistics_score),
            score_band_3(policy_score),
            inverse_score_band(fta_ratio, 40, 70),
            inverse_score_band(alt_score, 35, 55),
            score_band_3(top1_share, 60, 40),
            score_band_3(hhi, 2500, 1500),
            inverse_score_band(country_cnt, 4, 7),
            inverse_score_band(region_cnt, 2, 4),
        ]
    })

    st.markdown("#### 1) 현재 취약점 진단")
    st.dataframe(diag_df, use_container_width=True)

    # -----------------------------
    # 2) 대응 우선순위 자동 추천
    # -----------------------------
    def nz(v):
        return 0 if pd.isna(v) else float(v)

    action_need = {
        "공급국 다변화": 0.35*nz(top1_share) + 0.35*min(nz(hhi)/35, 100) + 0.20*(100-min(nz(country_cnt)*12.5, 100)) + 0.10*(100-nz(alt_score)),
        "FTA 활용 확대": 0.60*(100-nz(fta_ratio)) + 0.40*(100-nz(alt_score)),
        "상위 1국 의존도 축소": 0.60*nz(top1_share) + 0.40*min(nz(hhi)/35, 100),
        "물류 안전재고·운송다변화": 0.75*nz(logistics_score) + 0.25*(100-min(nz(region_cnt)*20, 100)),
        "정책/규제 대응력 강화": 1.00*nz(policy_score),
        "가격변동 대응(장기계약/헤지)": 1.00*nz(price_score),
    }

    reason_map = {
        "공급국 다변화": f"상위1국의존도({safe_fmt(top1_share)}%)·HHI({safe_fmt(hhi,0)})·수입국수({safe_fmt(country_cnt,0)}) 기준 공급집중 완화 필요",
        "FTA 활용 확대": f"FTA 비중({safe_fmt(fta_ratio)}%) 및 대체조달가능성({safe_fmt(alt_score)}점) 기준 조달 유연성 보완 필요",
        "상위 1국 의존도 축소": f"상위1국의존도({safe_fmt(top1_share)}%)와 집중도(HHI {safe_fmt(hhi,0)}) 기준 단일국 의존 완화 필요",
        "물류 안전재고·운송다변화": f"물류리스크점수({safe_fmt(logistics_score)}점)와 지역권수({safe_fmt(region_cnt,0)}) 기준 운송 병목 대응 필요",
        "정책/규제 대응력 강화": f"정책이벤트리스크점수({safe_fmt(policy_score)}점) 기준 정책 충격 대응 필요",
        "가격변동 대응(장기계약/헤지)": f"가격리스크점수({safe_fmt(price_score)}점) 기준 가격 변동성 완화 필요",
    }

    axis_map = {
        "공급국 다변화": "수급리스크",
        "FTA 활용 확대": "수급리스크 / 대체조달",
        "상위 1국 의존도 축소": "수급리스크",
        "물류 안전재고·운송다변화": "물류리스크",
        "정책/규제 대응력 강화": "정책이벤트리스크",
        "가격변동 대응(장기계약/헤지)": "가격리스크",
    }

    effect_map = {
        "공급국 다변화": "공급선 집중 완화, 수입국 수 확대",
        "FTA 활용 확대": "조달 유연성 확대, 관세·통관 측면 보완",
        "상위 1국 의존도 축소": "단일국 충격 노출 완화",
        "물류 안전재고·운송다변화": "운송 병목 및 지연 충격 완화",
        "정책/규제 대응력 강화": "제재·규제 이벤트 충격 완화",
        "가격변동 대응(장기계약/헤지)": "환율·원자재 가격 변동성 흡수",
    }

    rec_df = pd.DataFrame({
        "대응전략": list(action_need.keys()),
        "필요성점수": list(action_need.values())
    }).sort_values("필요성점수", ascending=False)

    rec_df["주요영향축"] = rec_df["대응전략"].map(axis_map)
    rec_df["기대효과"] = rec_df["대응전략"].map(effect_map)
    rec_df["근거"] = rec_df["대응전략"].map(reason_map)

    def priority_label(v):
        if v >= 75:
            return "최우선"
        elif v >= 60:
            return "우선"
        elif v >= 45:
            return "검토"
        else:
            return "후순위"

    rec_df["추천수준"] = rec_df["필요성점수"].apply(priority_label)

    st.markdown("#### 2) 대응 우선순위 자동 추천")
    st.dataframe(rec_df, use_container_width=True)

    # -----------------------------
    # 3) 시뮬레이션
    # -----------------------------
    st.markdown("#### 3) 대응 액션 설정")
    st.caption("""
각 슬라이더는 대응 강도를 의미하며, 조정 시 아래 리스크축/지표에 반영됩니다.

- 공급국 다변화 강화 → 수급리스크축 / 공급집중(HHI), 수입국수, 상위1국 의존 구조 완화 의미
- FTA 활용 확대 → 수급리스크축 및 대체조달 유연성 / fta_ratio 개선 의미
- 상위 1국 의존도 축소 → 수급리스크축 / 단일국 집중 완화 의미
- 물류 안전재고·운송다변화 → 물류리스크축 / GSCPI 기반 병목 노출 완화 의미
- 정책/규제 대응력 강화 → 정책이벤트리스크축 / 정책 이벤트 충격 흡수 의미
- 가격변동 대응(장기계약/헤지) → 가격리스크축 / 환율·원자재 가격 변동성 완화 의미

※ 현재 파일 구조상 일부 구조지표(HHI, 상위1국의존도, fta_ratio)는 직접 재계산 시뮬레이션까지는 하지 않고,
관련 리스크축 완화효과로 반영하는 참고형 모형입니다.
""")

    col1, col2 = st.columns(2)
    with col1:
        diversify = st.slider("공급국 다변화 강화", 0, 30, 10, help="영향축: 수급리스크 / 관련지표 의미: HHI·수입국수·공급집중 완화")
        expand_fta = st.slider("FTA 활용 확대", 0, 30, 10, help="영향축: 수급리스크 / 관련지표 의미: fta_ratio 개선")
        reduce_top1 = st.slider("상위 1국 의존도 축소", 0, 30, 10, help="영향축: 수급리스크 / 관련지표 의미: 상위1국 의존 구조 완화")
    with col2:
        logistics_buffer = st.slider("물류 안전재고·운송다변화", 0, 30, 10, help="영향축: 물류리스크 / 관련지표 의미: GSCPI 기반 물류병목 완화")
        policy_buffer = st.slider("정책/규제 대응력 강화", 0, 30, 10, help="영향축: 정책이벤트리스크 / 관련지표 의미: 정책 이벤트 충격 흡수")
        pricing_hedge = st.slider("가격변동 대응(장기계약/헤지)", 0, 30, 10, help="영향축: 가격리스크 / 관련지표 의미: 환율·원자재 가격 변동성 완화")

    sim_price = price_score * (1 - (pricing_hedge * 0.9) / 100) if pd.notna(price_score) else np.nan
    sim_supply = supply_score * (1 - (diversify * 0.5 + expand_fta * 0.2 + reduce_top1 * 0.5) / 100) if pd.notna(supply_score) else np.nan
    sim_logistics = logistics_score * (1 - (logistics_buffer * 0.9) / 100) if pd.notna(logistics_score) else np.nan
    sim_policy = policy_score * (1 - (policy_buffer * 0.9) / 100) if pd.notna(policy_score) else np.nan

    weights = None
    if entropy is not None and {"단계", "변수명", "가중치"}.issubset(entropy.columns):
        wf = entropy[entropy["단계"] == "PANEL_FINAL"].copy()
        if not wf.empty:
            weights = wf.set_index("변수명")["가중치"].to_dict()

    base_score = row["최종위험점수"] if "최종위험점수" in row.index else np.nan

    if weights:
        sim_raw = 0
        used = 0
        for c, v in {
            "가격리스크점수": sim_price,
            "수급리스크점수": sim_supply,
            "물류리스크점수": sim_logistics,
            "정책이벤트리스크점수": sim_policy
        }.items():
            if c in weights and pd.notna(v):
                sim_raw += v * weights[c]
                used += weights[c]
        sim_score = sim_raw / used if used > 0 else np.nan
    else:
        vals = [v for v in [sim_price, sim_supply, sim_logistics, sim_policy] if pd.notna(v)]
        sim_score = np.mean(vals) if vals else np.nan

    action_map_df = pd.DataFrame({
        "대응액션": [
            "공급국 다변화 강화",
            "FTA 활용 확대",
            "상위 1국 의존도 축소",
            "물류 안전재고·운송다변화",
            "정책/규제 대응력 강화",
            "가격변동 대응(장기계약/헤지)"
        ],
        "주요 영향 리스크축": [
            "수급리스크",
            "수급리스크 / 대체조달",
            "수급리스크",
            "물류리스크",
            "정책이벤트리스크",
            "가격리스크"
        ],
        "관련 지표 의미": [
            "HHI·수입국수·공급집중 구조",
            "fta_ratio·조달 유연성",
            "상위1국 의존도·집중도",
            "GSCPI 기반 병목 노출",
            "정책이벤트 충격",
            "환율·원자재 가격 변동성"
        ]
    })

    st.markdown("#### 액션별 영향 매핑")
    st.dataframe(action_map_df, use_container_width=True)

    st.markdown("#### 시뮬레이션 결과")
    c1, c2, c3 = st.columns(3)
    c1.metric("현재 최종위험점수", "-" if pd.isna(base_score) else f"{base_score:.2f}")
    c2.metric("개선 후 점수(참고)", "-" if pd.isna(sim_score) else f"{sim_score:.2f}")
    c3.metric("개선 폭", "-" if pd.isna(base_score) or pd.isna(sim_score) else f"{(base_score - sim_score):.2f}")

    compare_df = pd.DataFrame({
        "리스크축": ["가격", "수급", "물류", "정책"],
        "현재": [price_score, supply_score, logistics_score, policy_score],
        "개선후": [sim_price, sim_supply, sim_logistics, sim_policy]
    })

    fig = px.bar(
        compare_df.melt(id_vars="리스크축", var_name="구분", value_name="점수"),
        x="리스크축",
        y="점수",
        color="구분",
        barmode="group",
        title=f"{selected_chain} / {selected_month} 대응 전후 4대 리스크 비교"
    )
    st.plotly_chart(fig, use_container_width=True)

    if pd.notna(base_score) and pd.notna(sim_score):
        if base_score - sim_score >= 10:
            st.success("대응 효과가 비교적 크게 나타나는 시나리오입니다.")
        elif base_score - sim_score >= 5:
            st.info("중간 수준의 완화 효과가 예상됩니다.")
        else:
            st.warning("단기 완화 효과는 제한적입니다. 구조적 공급선 조정이 추가로 필요할 수 있습니다.")

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
elif menu == "7. 대체국 추천 시스템":
    st.subheader("대체국 추천 시스템")
    st.info(f"""
이 메뉴는 선택 기준월을 바탕으로 대체조달 후보국을 비교하는 화면입니다.

활용 방법
1) 안정성점수와 최종보정점수를 함께 봅니다.
2) FTA 여부와 상위공급국 여부를 고려해 대체조달 유연성을 판단합니다.
3) 추천점수가 높은 국가를 우선 검토하되, 실제 계약·품질·규격 조건과 함께 해석합니다.

참고
- 기본 기준월은 업로드된 데이터의 최신월인 **{latest_month if latest_month else '-'}** 입니다.
- 즉, 시스템 현재 날짜가 아니라 **업로드 데이터 기준 최신 관측월 추천**입니다.
- 데이터 범위: **{min_month if min_month else '-'} ~ {max_month if max_month else '-'}**
""")

    if country is None:
        st.warning("COUNTRY_MONTHLY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(country["체인구분"].dropna().unique().tolist()) if "체인구분" in country.columns else []
    if not chain_list:
        st.warning("체인구분 컬럼이 없습니다.")
        st.stop()

    selected_chain = st.selectbox("체인 선택", chain_list, key="alt_chain")

    month_list_country = sorted(country["연월"].dropna().unique().tolist()) if "연월" in country.columns else []
    selected_month = latest_month
    if month_list_country:
        default_idx = len(month_list_country) - 1
        if latest_month in month_list_country:
            default_idx = month_list_country.index(latest_month)
        selected_month = st.selectbox(
            "기준월 선택",
            month_list_country,
            index=default_idx,
            key="alt_month"
        )

    cur = country[(country["체인구분"] == selected_chain) & (country["연월"] == selected_month)].copy()

    if cur.empty:
        st.warning("선택 기준월 국가 데이터가 없습니다.")
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

    plot_df = cur.copy()
    agg_dict = {}
    for c in ["국가명", "지역권", "FTA여부", "상위공급국여부"]:
        if c in plot_df.columns:
            agg_dict[c] = "first"
    for c in ["기본평가점수", "최종보정점수", "안정성점수", "추천점수"]:
        if c in plot_df.columns:
            agg_dict[c] = "mean"

    if "국가코드" in plot_df.columns:
        plot_df = plot_df.groupby("국가코드", as_index=False).agg(agg_dict)
        plot_df = plot_df.sort_values("추천점수", ascending=False)

    show_cols = [c for c in [
        "국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부",
        "기본평가점수", "최종보정점수", "안정성점수", "추천점수"
    ] if c in plot_df.columns]
    st.dataframe(plot_df[show_cols].head(15), use_container_width=True)

    if "국가코드" in plot_df.columns and "추천점수" in plot_df.columns:
        fig = px.bar(
            plot_df.head(10),
            x="국가코드",
            y="추천점수",
            color="추천점수",
            color_continuous_scale="Blues",
            title=f"{selected_chain} / {selected_month} 대체국 추천 상위 10"
        )
        st.plotly_chart(fig, use_container_width=True)

    st.caption("※ 추천점수는 안정성, FTA 여부, 상위공급국 여부, 거래규모를 결합한 실무형 참고 점수입니다.")

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.subheader("원천데이터 탐색 / 다운로드")
    st.info("""
이 메뉴는 분석에 사용된 원천 및 중간가공 데이터를 직접 확인하는 화면입니다.

활용 방법
1) ITEM_INFO, DATA_INFO, COUNTRY_MONTHLY 등 주요 시트를 직접 확인합니다.
2) 원천값과 가공값을 비교해 데이터 신뢰성을 점검합니다.
3) 필요한 시트는 CSV로 다운로드해 별도 검토할 수 있습니다.
""")

    data_options = {
        "ITEM_INFO": item_info,
        "DATA_INFO": data_info,
        "RISK_MASTER": risk_master,
        "RISK_FALLBACK": risk_fallback,
        "MARKET_INDEX": market,
        "GSCPI_INDEX": gscpi,
        "TPU_INDEX": tpu,
        "HS_MONTHLY_SUMMARY": hs_summary,
        "COUNTRY_MONTHLY": country,
        "PANEL_MONTHLY": panel,
        "ALERT_RESULT": alert,
        "납산배터리 raw": leadacid_raw,
        "리튬이온배터리 raw": lithium_raw,
    }

    available_names = [k for k, v in data_options.items() if v is not None]
    selected_name = st.selectbox("탐색할 데이터 선택", available_names)
    selected_df = data_options[selected_name]

    st.markdown(f"#### {selected_name} 미리보기")

    if selected_name == "ITEM_INFO":
        st.dataframe(selected_df.iloc[:200, :10], use_container_width=True)
    else:
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
    st.info("""
이 메뉴는 앱의 계산값과 연결 구조가 정상인지 점검하는 검증 화면입니다.

활용 방법
1) 시트 존재 여부, 연월 파싱, 점수 범위를 자동 점검합니다.
2) OTH 처리, panel-alert 연결, GSCPI 컬럼명 등을 확인합니다.
3) 방법론과 가중치 시트를 함께 보며 계산 구조를 검토합니다.
""")

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
    map_df = pd.DataFrame({
        "논리명": list(sheet_map.keys()),
        "실사용 시트": list(sheet_map.values())
    })
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

    st.markdown("#### 6) 시스템 활용/갱신 방향")
    st.info("""
이 시스템은 2021~2025년 데이터를 기반으로 구축되었지만,
월별 신규 데이터를 동일 포맷으로 추가하면 같은 방식으로 계속 갱신 가능한 확장형 구조입니다.

활용 방향
- 2026년 이후: 관세청 수입데이터, 시장가격지표, GSCPI, TPU를 월별 추가 적재하여 점수 자동 갱신
- 2020년 이전: 데이터 확보 시 과거 위기국면 비교, 장기추세 분석, 구조변화 분석에 확장 가능
- 즉, 일회성 결과물이 아니라 상시 모니터링형 공급망 리스크 관리 체계로 발전 가능
""")

    st.markdown("#### 7) 검증 포함 엑셀 내보내기")
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
