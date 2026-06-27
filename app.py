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

    sub = panel[panel["
