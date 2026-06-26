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
    10월 깨짐 방지용 연월 파서.
    예:
    - 2021.01 -> 2021-01
    - 2021.1  -> 2021-10 으로 해석
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

        # YYYY-MM / YYYY/M / YYYY.MM
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

        # 일반 날짜형 파싱
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

        # 2021.0 같은 형태는 1월로 처리
        if abs(frac) < 1e-12:
            return f"{y:04d}-01"

        frac_str = f"{frac:.10f}".split(".")[1].rstrip("0")

        try:
            # 2021.1 -> 10월
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

# =========================================================
# 데이터 로드
# =========================================================
@st.cache_data(show_spinner=False)
def load_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    # 현재 네 최종 파일 구조 기준
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

    # GSCPI 컬럼명 점검
    gscpi = data.get("gscpi")
    if gscpi is not None:
        cols = list(gscpi.columns)
        if "GSCPI" in cols and "GSCPI_NORM" in cols:
            add("PASS", "GSCPI 컬럼명", "GSCPI / GSCPI_NORM 정상")
        else:
            add("WARN", "GSCPI 컬럼명", f"현재 컬럼: {cols}")

    # 연월 파싱
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

    # 범위 점검
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

    # fta_ratio
    if panel is not None and "fta_ratio" in panel.columns:
        s = safe_numeric(panel["fta_ratio"])
        bad = ((s < 0) | (s > 100)).sum()
        if bad == 0:
            add("PASS", "fta_ratio 범위", "0~100 정상")
        else:
            add("FAIL", "fta_ratio 범위", f"0~100 이탈 {bad}건")

    # OTH 점검
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

    # panel-alert 연결
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

    # 체인별비교표 우선관리대상
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
    c3.metric("검증 WARN", int((check_df["레벨"] == "WARN").sum()))
    c4.metric("검증 FAIL", int((check_df["레벨"] == "FAIL").sum()))

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
                x="연월",
                y="최종위험점수",
                color="체인구분",
                markers=True,
                title="체인별 최종위험점수 추이"
            )
            st.plotly_chart(fig, use_container_width=True)

        risk_cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in latest_panel.columns]
        if risk_cols:
            bar_df = latest_panel[["체인구분"] + risk_cols].melt(id_vars="체인구분", var_name="리스크축", value_name="점수")
            fig2 = px.bar(
                bar_df,
                x="리스크축",
                y="점수",
                color="체인구분",
                barmode="group",
                title=f"{latest_month} 4대 리스크 비교"
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

        # 리튬이온배터리군 설명
        if "체인구분" in chain_compare.columns and "우선관리대상_비중" in chain_compare.columns:
            lithium_rows = chain_compare[chain_compare["체인구분"].astype(str).str.contains("리튬", na=False)]
            if not lithium_rows.empty:
                val = lithium_rows["우선관리대상_비중"].iloc[0]
                if pd.notna(val) and float(val) == 0:
                    st.caption("""
리튬이온배터리군의 우선관리대상 비중이 0%라는 것은 위험이 없다는 뜻이 아니라,
분석기간 내 '고위험'과 '저대체조달'이 동시에 강하게 충족된 월이 없었다는 의미입니다.
즉 위험 상승 구간은 있었지만, 우선관리 판정 규칙상 동시충족 빈도가 낮았던 것입니다.
""")

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

    candidate_cols = [c for c in [
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
        "경보점수기초", "국가보정합계", "환율정규화", "납가격정규화",
        "리튬가격정규화", "니켈가격정규화", "GSCPI_Norm", "정책이벤트_raw", "최종위험점수"
    ] if c in sub.columns]

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
# 6. 기업 대응 시뮬레이터
# =========================================================
elif menu == "6. 기업 대응 시뮬레이터":
    st.subheader("기업 대응 시뮬레이터")

    st.info("""
이 기능은 최신월 기준으로 기업이 취할 수 있는 대응조치를 가정했을 때,
4대 리스크 축과 최종위험점수가 얼마나 완화될 수 있는지 시뮬레이션하는 참고 도구입니다.
정확한 예측모형이 아니라, 대응 우선순위를 비교하기 위한 실무형 실험 기능입니다.
""")

    if panel is None:
        st.warning("PANEL_MONTHLY 데이터가 없습니다.")
        st.stop()

    chain_list = sorted(panel["체인구분"].dropna().unique().tolist())
    selected_chain = st.selectbox("체인 선택", chain_list, key="sim_chain")

    base = panel[(panel["체인구분"] == selected_chain) & (panel["연월"] == latest_month)].copy()
    if base.empty:
        st.warning("최신월 데이터가 없습니다.")
        st.stop()

    row = base.iloc[0].copy()

    st.markdown("#### 대응 액션 설정")

    col1, col2 = st.columns(2)
    with col1:
        diversify = st.slider("공급국 다변화 강화", 0, 30, 10, help="수급리스크 완화에 주로 반영")
        expand_fta = st.slider("FTA 활용 확대", 0, 30, 10, help="수급리스크 및 대체조달 측면 개선")
        reduce_top1 = st.slider("상위 1국 의존도 축소", 0, 30, 10, help="수급리스크 완화에 반영")
    with col2:
        logistics_buffer = st.slider("물류 안전재고·운송다변화", 0, 30, 10, help="물류리스크 완화")
        policy_buffer = st.slider("정책/규제 대응력 강화", 0, 30, 10, help="정책리스크 완화")
        pricing_hedge = st.slider("가격변동 대응(장기계약/헤지)", 0, 30, 10, help="가격리스크 완화")

    base_price = row["가격리스크점수"] if "가격리스크점수" in row.index else np.nan
    base_supply = row["수급리스크점수"] if "수급리스크점수" in row.index else np.nan
    base_logistics = row["물류리스크점수"] if "물류리스크점수" in row.index else np.nan
    base_policy = row["정책이벤트리스크점수"] if "정책이벤트리스크점수" in row.index else np.nan

    sim_price = base_price * (1 - (pricing_hedge * 0.9) / 100) if pd.notna(base_price) else np.nan
    sim_supply = base_supply * (1 - (diversify * 0.5 + expand_fta * 0.2 + reduce_top1 * 0.5) / 100) if pd.notna(base_supply) else np.nan
    sim_logistics = base_logistics * (1 - (logistics_buffer * 0.9) / 100) if pd.notna(base_logistics) else np.nan
    sim_policy = base_policy * (1 - (policy_buffer * 0.9) / 100) if pd.notna(base_policy) else np.nan

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

    st.markdown("#### 시뮬레이션 결과")

    c1, c2, c3 = st.columns(3)
    c1.metric("현재 최종위험점수", "-" if pd.isna(base_score) else f"{base_score:.2f}")
    c2.metric("개선 후 점수(참고)", "-" if pd.isna(sim_score) else f"{sim_score:.2f}")
    c3.metric("개선 폭", "-" if pd.isna(base_score) or pd.isna(sim_score) else f"{(base_score - sim_score):.2f}")

    compare_df = pd.DataFrame({
        "리스크축": ["가격", "수급", "물류", "정책"],
        "현재": [base_price, base_supply, base_logistics, base_policy],
        "개선후": [sim_price, sim_supply, sim_logistics, sim_policy]
    })

    fig = px.bar(
        compare_df.melt(id_vars="리스크축", var_name="구분", value_name="점수"),
        x="리스크축",
        y="점수",
        color="구분",
        barmode="group",
        title=f"{selected_chain} 대응 전후 4대 리스크 비교"
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

    if country is None:
        st.warning("COUNTRY_MONTHLY 데이터가 없습니다.")
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

    # 차트용 국가별 1행 집계
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
            title=f"{selected_chain} 대체국 추천 상위 10"
        )
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
