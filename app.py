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

def pick_existing_sheet(sheet_names, candidates):
    for s in candidates:
        if s in sheet_names:
            return s
    return None

def safe_ym(x, year=None, month=None):
    # 날짜 비교/정렬은 반드시 이 표준화 연월만 사용
    if year is not None and month is not None and pd.notna(year) and pd.notna(month):
        try:
            return f"{int(year):04d}-{int(month):02d}"
        except Exception:
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
                except Exception:
                    pass
        try:
            dt = pd.to_datetime(s)
            return f"{dt.year:04d}-{dt.month:02d}"
        except Exception:
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
        except Exception:
            pass

        try:
            m = int(round((float(x) - y) * 100))
            if 1 <= m <= 12:
                return f"{y:04d}-{m:02d}"
        except Exception:
            pass

    return str(x)

def ensure_ym_column(df):
    df = df.copy()
    cols = list(df.columns)

    if "연월" in cols:
        if "연도" in cols and "월" in cols:
            df["연월_표준"] = df.apply(lambda r: safe_ym(r["연월"], r["연도"], r["월"]), axis=1)
        else:
            df["연월_표준"] = df["연월"].apply(safe_ym)
        df["연월"] = df["연월_표준"]
        return df

    if "연도" in cols and "월" in cols:
        df["연월_표준"] = df.apply(lambda r: safe_ym(None, r["연도"], r["월"]), axis=1)
        df["연월"] = df["연월_표준"]
        return df

    if "연" in cols and "월" in cols:
        df["연월_표준"] = df.apply(lambda r: safe_ym(None, r["연"], r["월"]), axis=1)
        df["연월"] = df["연월_표준"]
        return df

    return df

def fmt_num(v, digits=2):
    if pd.isna(v):
        return "-"
    try:
        return f"{float(v):,.{digits}f}"
    except Exception:
        return str(v)

def fmt_pct(v, digits=2):
    if pd.isna(v):
        return "-"
    return f"{float(v):.{digits}f}%"

def to_download_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            if df is None or not isinstance(df, pd.DataFrame):
                continue
            temp = df.copy()
            for c in temp.columns:
                if pd.api.types.is_datetime64_any_dtype(temp[c]):
                    temp[c] = temp[c].dt.strftime("%Y-%m")
            temp.to_excel(writer, sheet_name=sname[:31], index=False)
    return output.getvalue()

def get_chain_list(panel):
    if panel is None or "체인구분" not in panel.columns:
        return []
    vals = panel["체인구분"].dropna().astype(str).unique().tolist()
    return sorted(vals)

def get_month_list(df):
    if df is None or "연월" not in df.columns:
        return []
    vals = df["연월"].dropna().astype(str).unique().tolist()
    return sorted(vals)

def find_col(df, candidates):
    if df is None:
        return None
    cols = list(df.columns)
    lower_map = {str(c).lower(): c for c in cols}
    for c in candidates:
        if c in cols:
            return c
        if str(c).lower() in lower_map:
            return lower_map[str(c).lower()]
    return None

def prepare_aliases(data):
    # TPU alias
    tpu = data.get("tpu")
    if tpu is not None:
        if "TPU_INDEX_NORM" not in tpu.columns and "정책이벤트정규화" in tpu.columns:
            tpu["TPU_INDEX_NORM"] = tpu["정책이벤트정규화"]
        if "정책이벤트정규화" not in tpu.columns and "TPU_INDEX_NORM" in tpu.columns:
            tpu["정책이벤트정규화"] = tpu["TPU_INDEX_NORM"]
        data["tpu"] = tpu

    gscpi = data.get("gscpi")
    if gscpi is not None:
        if "GSCPI_Norm" not in gscpi.columns and "GSCPI_NORM" in gscpi.columns:
            gscpi["GSCPI_Norm"] = gscpi["GSCPI_NORM"]
        if "GSCPI_NORM" not in gscpi.columns and "GSCPI_Norm" in gscpi.columns:
            gscpi["GSCPI_NORM"] = gscpi["GSCPI_Norm"]
        data["gscpi"] = gscpi

    panel = data.get("panel")
    if panel is not None:
        if "정책이벤트리스크점수" not in panel.columns and "TPU_INDEX_NORM" in panel.columns:
            panel["정책이벤트리스크점수"] = panel["TPU_INDEX_NORM"]
        data["panel"] = panel

    return data

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
        "norm_audit": pick_existing_sheet(sheet_names, ["NOMALIZATION_AUDIT", "NORMALIZATION_AUDIT"]),
        "method": pick_existing_sheet(sheet_names, ["METHOD_GUIDE"]),
        "signal_scope": pick_existing_sheet(sheet_names, ["SIGNAL_CANDIDATE_SCOPE"]),
        "signal_base": pick_existing_sheet(sheet_names, ["SIGNAL_BASE"]),
        "signal_lag": pick_existing_sheet(sheet_names, ["SIGNAL_LAG_TABLE"]),
        "lead_detail": pick_existing_sheet(sheet_names, ["LEAD_SIGNAL_LAG_DETAIL"]),
        "lead_compare": pick_existing_sheet(sheet_names, ["LEAD_SIGNAL_LAG_COMPARE"]),
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
            "보정사유", "지역권", "품목명", "품목군", "배터리유형", "비고", "선정이유"
        ]:
            df = normalize_text_col(df, txt_col)

        data[key] = df

    data = prepare_aliases(data)
    return sheet_names, sheet_map, data

# =========================================================
# 검증 로직
# =========================================================
def run_checks(sheet_names, sheet_map, data):
    results = []

    def add(level, item, detail):
        results.append({"레벨": level, "점검항목": item, "상세": detail})

    required_keys = [
        "country", "panel", "alert", "chain_compare", "entropy",
        "norm_check", "method", "signal_scope", "signal_base", "signal_lag",
        "lead_detail", "lead_compare"
    ]
    for k in required_keys:
        if data.get(k) is None:
            add("FAIL", f"필수 시트 누락", f"{k} 시트를 찾지 못했습니다.")
        else:
            add("PASS", f"필수 시트 존재", f"{sheet_map[k]} 사용")

    panel = data.get("panel")
    alert = data.get("alert")
    tpu = data.get("tpu")
    gscpi = data.get("gscpi")
    lead_detail = data.get("lead_detail")
    lead_compare = data.get("lead_compare")
    signal_base = data.get("signal_base")

    if tpu is not None:
        if {"TPU_INDEX", "TPU_INDEX_NORM"}.issubset(tpu.columns):
            add("PASS", "TPU_INDEX 컬럼", "TPU_INDEX / TPU_INDEX_NORM 존재")
        else:
            add("FAIL", "TPU_INDEX 컬럼", "TPU_INDEX 또는 TPU_INDEX_NORM 누락")

    if gscpi is not None:
        if {"GSCPI", "GSCPI_Norm"}.issubset(gscpi.columns) or {"GSCPI", "GSCPI_NORM"}.issubset(gscpi.columns):
            add("PASS", "GSCPI 컬럼", "GSCPI / GSCPI_Norm 존재")
        else:
            add("FAIL", "GSCPI 컬럼", "GSCPI 또는 GSCPI_Norm 누락")

    if panel is not None:
        required_cols = [
            "연월", "체인구분", "HHI", "상위1국의존도", "경보점수기초", "국가보정합계",
            "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
            "최종위험점수", "최종경보등급", "fta_ratio", "지역권수", "TPU_INDEX_NORM"
        ]
        missing = [c for c in required_cols if c not in panel.columns]
        if missing:
            add("FAIL", "PANEL_MONTHLY 필수 컬럼", f"누락: {missing}")
        else:
            add("PASS", "PANEL_MONTHLY 필수 컬럼", "핵심 컬럼 존재")

        if "CV" in panel.columns:
            add("WARN", "CV 제거 여부", "PANEL_MONTHLY에 CV 컬럼이 남아 있습니다.")
        else:
            add("PASS", "CV 제거 여부", "PANEL_MONTHLY에 CV 컬럼 없음")

        if {"TPU_INDEX_NORM", "정책이벤트리스크점수"}.issubset(panel.columns):
            x = safe_numeric(panel["TPU_INDEX_NORM"]).round(4)
            y = safe_numeric(panel["정책이벤트리스크점수"]).round(4)
            if x.fillna(-9999).equals(y.fillna(-9999)):
                add("PASS", "정책축 값 일치", "TPU_INDEX_NORM == 정책이벤트리스크점수")
            else:
                add("WARN", "정책축 값 일치", "TPU_INDEX_NORM 과 정책이벤트리스크점수 일부 불일치")

    if panel is not None and alert is not None:
        merge_cols = [
            "가격리스크점수", "수급리스크점수", "물류리스크점수",
            "정책이벤트리스크점수", "최종위험점수", "최종경보등급"
        ]
        try:
            merged = alert.merge(
                panel[["연월", "체인구분"] + merge_cols],
                on=["연월", "체인구분"],
                how="left",
                suffixes=("_alert", "_panel")
            )
            mismatches = 0
            for c in merge_cols:
                ca, cp = f"{c}_alert", f"{c}_panel"
                if c == "최종경보등급":
                    mismatches += int((merged[ca].fillna("") != merged[cp].fillna("")).sum())
                else:
                    a = safe_numeric(merged[ca]).round(4)
                    p = safe_numeric(merged[cp]).round(4)
                    mismatches += int((a.fillna(-9999) != p.fillna(-9999)).sum())
            if mismatches == 0:
                add("PASS", "ALERT_RESULT vs PANEL_MONTHLY", "핵심 점수/등급 일치")
            else:
                add("WARN", "ALERT_RESULT vs PANEL_MONTHLY", f"불일치 셀 수: {mismatches}")
        except Exception as e:
            add("WARN", "ALERT_RESULT vs PANEL_MONTHLY", f"검증 중 오류: {e}")

    if lead_detail is not None:
        if "lag개월" in lead_detail.columns:
            lag_vals = sorted(pd.Series(lead_detail["lag개월"]).dropna().astype(int).unique().tolist())
            if lag_vals == [1, 2, 3, 4, 5, 6]:
                add("PASS", "LEAD_SIGNAL_LAG_DETAIL lag 범위", "1~6개월 정상")
            else:
                add("WARN", "LEAD_SIGNAL_LAG_DETAIL lag 범위", f"현재 값: {lag_vals}")
        else:
            add("FAIL", "LEAD_SIGNAL_LAG_DETAIL lag 컬럼", "lag개월 컬럼 누락")

    if lead_compare is not None:
        need = ["지표", "최적lag개월", "최적lag상관", "최적lag절대상관", "최종해석"]
        missing = [c for c in need if c not in lead_compare.columns]
        if missing:
            add("FAIL", "LEAD_SIGNAL_LAG_COMPARE 필수 컬럼", f"누락: {missing}")
        else:
            add("PASS", "LEAD_SIGNAL_LAG_COMPARE 필수 컬럼", "정상")

    if signal_base is not None:
        if "최종위험점수_V2" in signal_base.columns:
            add("PASS", "SIGNAL_BASE 종속변수", "최종위험점수_V2 존재")
        else:
            add("FAIL", "SIGNAL_BASE 종속변수", "최종위험점수_V2 누락")

    return pd.DataFrame(results)

# =========================================================
# 화면 표시용 공통 준비
# =========================================================
def latest_month_snapshot(panel):
    if panel is None or panel.empty:
        return None, None
    latest = sorted(panel["연월"].dropna().astype(str).unique().tolist())[-1]
    return latest, panel[panel["연월"] == latest].copy()

def get_entropy_weights(entropy_df, stage_name):
    if entropy_df is None or entropy_df.empty:
        return {}
    sub = entropy_df[entropy_df["단계"] == stage_name].copy()
    if sub.empty:
        return {}
    return dict(zip(sub["변수명"], safe_numeric(sub["가중치"])))

def build_priority_comment(row):
    comments = []
    if safe_numeric(pd.Series([row.get("최종위험점수", np.nan)])).iloc[0] >= 75:
        comments.append("최종위험 수준이 높아 즉시 점검이 필요합니다.")
    if safe_numeric(pd.Series([row.get("상위1국의존도", np.nan)])).iloc[0] >= 70:
        comments.append("상위 1개국 의존도가 높아 공급선 다변화 검토가 필요합니다.")
    if safe_numeric(pd.Series([row.get("fta_ratio", np.nan)])).iloc[0] <= 30:
        comments.append("FTA 활용 비중이 낮아 비용/통관 측면 개선 여지가 있습니다.")
    if safe_numeric(pd.Series([row.get("대체조달가능성_점수", np.nan)])).iloc[0] <= 30:
        comments.append("대체조달가능성이 낮아 대체국 발굴 우선순위가 높습니다.")
    if not comments:
        comments.append("현재 지표상 급격한 경보는 아니며 정기 모니터링이 적절합니다.")
    return " ".join(comments)

def show_df_download_button(df, label, filename):
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(label=label, data=csv, file_name=filename, mime="text/csv")

# =========================================================
# 상단
# =========================================================
st.title("망보는사람들 공급망 리스크 대시보드")
st.caption("최종 엑셀 구조 기준 · 날짜 비교/정렬은 표준화된 연월(YYYY-MM)만 사용")

uploaded_file = st.sidebar.file_uploader("최종 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file is None:
    st.info("최종 엑셀 파일(.xlsx)을 업로드하면 대시보드가 열립니다.")
    st.stop()

sheet_names, sheet_map, data = load_workbook(uploaded_file)

panel = data.get("panel")
alert = data.get("alert")
country = data.get("country")
chain_compare = data.get("chain_compare")
entropy = data.get("entropy")
norm_check = data.get("norm_check")
norm_audit = data.get("norm_audit")
method = data.get("method")
signal_scope = data.get("signal_scope")
signal_base = data.get("signal_base")
signal_lag = data.get("signal_lag")
lead_detail = data.get("lead_detail")
lead_compare = data.get("lead_compare")
hs_summary = data.get("hs_summary")
gscpi = data.get("gscpi")
tpu = data.get("tpu")
item_info = data.get("item_info")
data_info = data.get("data_info")

menu = st.sidebar.radio(
    "메뉴 선택",
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
    ],
)

# =========================================================
# 1. 종합 상황판
# =========================================================
if menu == "1. 종합 상황판":
    st.header("1. 종합 상황판")

    if panel is None or chain_compare is None:
        st.error("PANEL_MONTHLY 또는 체인별 비교표 시트가 없습니다.")
        st.stop()

    latest_month, latest_df = latest_month_snapshot(panel)
    st.subheader(f"최신 스냅샷: {latest_month}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("분석 체인 수", latest_df["체인구분"].nunique())
    c2.metric("평균 최종위험점수", fmt_num(safe_numeric(latest_df["최종위험점수"]).mean(), 2))
    c3.metric("최고 최종위험점수", fmt_num(safe_numeric(latest_df["최종위험점수"]).max(), 2))
    c4.metric("매우높음 체인 수", int((latest_df["최종경보등급"] == "매우높음").sum()))

    st.markdown("#### 최신월 체인별 요약")
    show_cols = [
        "체인구분", "가격리스크점수", "수급리스크점수", "물류리스크점수",
        "정책이벤트리스크점수", "최종위험점수", "최종경보등급",
        "상위1국의존도", "HHI", "fta_ratio", "지역권수"
    ]
    st.dataframe(latest_df[show_cols], use_container_width=True)

    st.markdown("#### 체인별 최종위험 추이")
    fig = px.line(
        panel.sort_values("연월"),
        x="연월",
        y="최종위험점수",
        color="체인구분",
        markers=True
    )
    fig.update_layout(height=420)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 4대 리스크 축 비교")
    axis_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    melt_df = latest_df.melt(id_vars=["체인구분"], value_vars=axis_cols, var_name="리스크축", value_name="점수")
    fig2 = px.bar(melt_df, x="리스크축", y="점수", color="체인구분", barmode="group")
    fig2.update_layout(height=420)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 체인별 비교표")
    st.dataframe(chain_compare, use_container_width=True)

    st.markdown("#### 리스크 산식 기준")
    risk_basis_df = pd.DataFrame({
        "리스크축": ["가격리스크", "수급리스크", "물류리스크", "정책이벤트리스크"],
        "주요 입력값": [
            "환율정규화, 납가격정규화, 리튬가격정규화, 니켈가격정규화",
            "HHI, 상위1국의존도 기반 경보점수기초 + 국가보정합계",
            "GSCPI_Norm",
            "TPU_INDEX_NORM"
        ]
    })
    st.dataframe(risk_basis_df, use_container_width=True)

# =========================================================
# 2. 체인별 심층 분석
# =========================================================
if menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chains = get_chain_list(panel)
    chain = st.selectbox("체인 선택", chains)
    sub = panel[panel["체인구분"] == chain].sort_values("연월").copy()

    latest_month = sorted(sub["연월"].dropna().astype(str).unique().tolist())[-1]
    latest_row = sub[sub["연월"] == latest_month].iloc[0]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("최신 최종위험점수", fmt_num(latest_row["최종위험점수"], 2), latest_row["최종경보등급"])
    c2.metric("상위1국의존도", fmt_pct(latest_row["상위1국의존도"], 2))
    c3.metric("HHI", fmt_num(latest_row["HHI"], 2))
    c4.metric("FTA 비중", fmt_pct(latest_row["fta_ratio"], 2))

    trend_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수"]
    melt = sub.melt(id_vars=["연월"], value_vars=trend_cols, var_name="지표", value_name="값")
    fig = px.line(melt, x="연월", y="값", color="지표", markers=True)
    fig.update_layout(height=460)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 수입 구조와 리스크 요약")
    display_cols = [
        "연월", "총수입금액", "총수입물량", "평균수입단가", "수입국수", "지역권수",
        "상위1국의존도", "상위3국집중도", "HHI", "경보점수기초", "국가보정합계",
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
        "최종위험점수", "최종경보등급"
    ]
    st.dataframe(sub[display_cols], use_container_width=True)

    if hs_summary is not None:
        st.markdown("#### HS 품목 월별 요약")
        hs_sub = hs_summary[hs_summary["체인구분"] == chain].sort_values(["연월", "HS코드"])
        hs_cols = [c for c in [
            "연월", "HS코드", "품목명", "총수입중량", "총수입금액", "평균수입단가",
            "상위공급국", "상위공급국비중", "수입국수"
        ] if c in hs_sub.columns]
        st.dataframe(hs_sub[hs_cols], use_container_width=True)

# =========================================================
# 3. 국가/공급선 상세 분석
# =========================================================
if menu == "3. 국가/공급선 상세 분석":
    st.header("3. 국가/공급선 상세 분석")

    if country is None:
        st.error("COUNTRY_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(country), key="country_chain")
    month_list = get_month_list(country[country["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=len(month_list)-1 if month_list else 0)

    sub = country[(country["체인구분"] == chain) & (country["연월"] == month)].copy()
    sub = sub.sort_values(["금액비중", "국가수입금액"], ascending=[False, False])

    c1, c2, c3 = st.columns(3)
    c1.metric("공급국 수", sub["국가명"].nunique())
    c2.metric("평균 최종보정점수", fmt_num(safe_numeric(sub["최종보정점수"]).mean(), 2))
    c3.metric("상위공급국 수", int((sub["상위공급국여부"] == "Y").sum()) if "상위공급국여부" in sub.columns else 0)

    topn = st.slider("상위 몇 개 국가를 볼지", 5, min(30, len(sub)) if len(sub) > 0 else 5, 10)
    top = sub.head(topn)

    fig = px.bar(
        top,
        x="국가명",
        y="금액비중",
        color="최종판정" if "최종판정" in top.columns else None,
        hover_data=["국가코드", "FTA여부", "최종보정점수"]
    )
    fig.update_layout(height=420)
    st.plotly_chart(fig, use_container_width=True)

    fig2 = px.scatter(
        sub,
        x="금액비중",
        y="최종보정점수",
        size="국가수입금액" if "국가수입금액" in sub.columns else None,
        color="FTA여부" if "FTA여부" in sub.columns else None,
        hover_name="국가명"
    )
    fig2.update_layout(height=420)
    st.plotly_chart(fig2, use_container_width=True)

    show_cols = [c for c in [
        "연월", "국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부",
        "국가수입금액", "금액비중", "기본평가점수", "공급국집중보정", "지역권집중보정",
        "HHI보정", "상위공급국보정", "FTA보정", "총보정점수", "최종보정점수", "최종판정", "비고"
    ] if c in sub.columns]
    st.dataframe(sub[show_cols], use_container_width=True)
    show_df_download_button(sub[show_cols], "국가 상세표 CSV 다운로드", f"country_detail_{chain}_{month}.csv")

# =========================================================
# 4. 충격 원인 추적
# =========================================================
if menu == "4. 충격 원인 추적":
    st.header("4. 충격 원인 추적")

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel), key="shock_chain")
    month_list = get_month_list(panel[panel["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=len(month_list)-1 if month_list else 0, key="shock_month")

    row = panel[(panel["체인구분"] == chain) & (panel["연월"] == month)].copy()
    if row.empty:
        st.warning("선택 조건에 해당하는 데이터가 없습니다.")
        st.stop()
    row = row.iloc[0]

    axis_df = pd.DataFrame({
        "리스크축": ["가격리스크", "수급리스크", "물류리스크", "정책이벤트리스크"],
        "점수": [
            row.get("가격리스크점수", np.nan),
            row.get("수급리스크점수", np.nan),
            row.get("물류리스크점수", np.nan),
            row.get("정책이벤트리스크점수", np.nan),
        ]
    }).sort_values("점수", ascending=False)

    c1, c2, c3 = st.columns(3)
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan), 2), row.get("최종경보등급", "-"))
    c2.metric("주요 원인축", axis_df.iloc[0]["리스크축"])
    c3.metric("상위1국의존도", fmt_pct(row.get("상위1국의존도", np.nan), 2))

    fig = px.bar(axis_df, x="리스크축", y="점수", color="리스크축")
    fig.update_layout(height=380, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 세부 원인 설명")
    reasons = []
    reasons.append(f"- 가격리스크점수: {fmt_num(row.get('가격리스크점수', np.nan), 2)}")
    reasons.append(f"- 수급리스크점수: {fmt_num(row.get('수급리스크점수', np.nan), 2)} (HHI {fmt_num(row.get('HHI', np.nan), 2)}, 상위1국의존도 {fmt_pct(row.get('상위1국의존도', np.nan), 2)}, 국가보정합계 {fmt_num(row.get('국가보정합계', np.nan), 2)})")
    reasons.append(f"- 물류리스크점수: {fmt_num(row.get('물류리스크점수', np.nan), 2)} (GSCPI_Norm 기준)")
    reasons.append(f"- 정책이벤트리스크점수: {fmt_num(row.get('정책이벤트리스크점수', np.nan), 2)} (TPU_INDEX_NORM 기준)")
    st.markdown("\n".join(reasons))

    if gscpi is not None:
        gsub = gscpi[gscpi["연월"] == month]
        if not gsub.empty:
            g = gsub.iloc[0]
            st.info(f"GSCPI 원천값: {fmt_num(g.get('GSCPI', np.nan), 2)} / 정규화값: {fmt_num(g.get('GSCPI_Norm', g.get('GSCPI_NORM', np.nan)), 2)}")

    if tpu is not None:
        tsub = tpu[tpu["연월"] == month]
        if not tsub.empty:
            t = tsub.iloc[0]
            narrative = t.get("서사배경", "-")
            st.info(
                f"TPU 원천값: {fmt_num(t.get('TPU_INDEX', np.nan), 2)} / "
                f"정규화값(TPU_INDEX_NORM): {fmt_num(t.get('TPU_INDEX_NORM', np.nan), 2)}"
            )
            st.markdown("#### 정책 이벤트 배경")
            st.write(narrative)

# =========================================================
# 5. 선행 신호 후보 탐지
# =========================================================
if menu == "5. 선행 신호 후보 탐지":
    st.header("5. 선행 신호 후보 탐지")
    st.caption("가공 점수 대신 원재료/원천 변수 중심으로 lag 1~6개월을 비교한 검증 결과")

    if signal_scope is None or lead_compare is None or lead_detail is None:
        st.error("SIGNAL / LEAD_SIGNAL 관련 시트가 없습니다.")
        st.stop()

    chain_options = sorted(lead_compare["체인구분"].dropna().astype(str).unique().tolist())
    chain = st.selectbox("체인 선택", chain_options, key="lead_chain")

    st.markdown("#### 1) 후보 변수 범위")
    scope_sub = signal_scope[
        (signal_scope["체인구분"].isin(["공통", chain])) &
        (signal_scope["포함여부"].isin(["Y", "N"]))
    ].copy()
    st.dataframe(scope_sub, use_container_width=True)

    st.markdown("#### 2) 변수별 최적 lag 요약")
    compare_sub = lead_compare[lead_compare["체인구분"] == chain].copy()
    compare_sub = compare_sub.sort_values("최적lag절대상관", ascending=False)
    st.dataframe(compare_sub, use_container_width=True)

    fig = px.bar(
        compare_sub,
        x="지표",
        y="최적lag상관",
        color="최종해석",
        hover_data=["최적lag개월", "최적lag절대상관", "최적lag관측치N"]
    )
    fig.update_layout(height=430)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 3) 지표별 lag 1~6 상세")
    metric = st.selectbox("지표 선택", compare_sub["지표"].tolist(), key="lead_metric")
    detail_sub = lead_detail[(lead_detail["체인구분"] == chain) & (lead_detail["지표"] == metric)].copy()
    st.dataframe(detail_sub, use_container_width=True)

    fig2 = px.line(detail_sub, x="lag개월", y="상관계수", markers=True)
    fig2.update_layout(height=380, xaxis=dict(dtick=1))
    st.plotly_chart(fig2, use_container_width=True)

    if signal_base is not None:
        st.markdown("#### 4) 검증 베이스 데이터 미리보기")
        base_cols = ["연월", "체인구분", "최종위험점수_V2", metric]
        base_cols = [c for c in base_cols if c in signal_base.columns]
        base_sub = signal_base[signal_base["체인구분"] == chain][base_cols].copy()
        st.dataframe(base_sub, use_container_width=True)

# =========================================================
# 6. 기업 대응 우선순위 추천 / 시뮬레이터
# =========================================================
if menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.header("6. 기업 대응 우선순위 추천 / 시뮬레이터")

    if panel is None or alert is None:
        st.error("PANEL_MONTHLY 또는 ALERT_RESULT 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel), key="prio_chain")
    month_list = get_month_list(panel[panel["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=len(month_list)-1 if month_list else 0, key="prio_month")

    panel_row = panel[(panel["체인구분"] == chain) & (panel["연월"] == month)].copy()
    alert_row = alert[(alert["체인구분"] == chain) & (alert["연월"] == month)].copy()

    if panel_row.empty:
        st.warning("선택한 조건의 데이터가 없습니다.")
        st.stop()

    row = panel_row.iloc[0].to_dict()
    if not alert_row.empty:
        row.update(alert_row.iloc[0].to_dict())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan), 2), row.get("최종경보등급", "-"))
    c2.metric("대체조달가능성", fmt_num(row.get("대체조달가능성_점수", np.nan), 2), row.get("대체조달가능성", "-"))
    c3.metric("보정사유", row.get("보정사유", "-"))
    c4.metric("우선관리대상", row.get("상대적_우선관리대상", "-"))

    st.markdown("#### 대응 우선순위 진단")
    diagnose_df = pd.DataFrame({
        "항목": ["상위1국의존도", "HHI", "FTA 비중", "지역권수", "수입국수"],
        "현재값": [
            row.get("상위1국의존도", np.nan),
            row.get("HHI", np.nan),
            row.get("fta_ratio", np.nan),
            row.get("지역권수", np.nan),
            row.get("수입국수", np.nan),
        ],
        "판정": [
            "취약" if safe_numeric(pd.Series([row.get("상위1국의존도", np.nan)])).iloc[0] >= 70 else "주의" if safe_numeric(pd.Series([row.get("상위1국의존도", np.nan)])).iloc[0] >= 50 else "양호",
            "취약" if safe_numeric(pd.Series([row.get("HHI", np.nan)])).iloc[0] >= 5000 else "주의" if safe_numeric(pd.Series([row.get("HHI", np.nan)])).iloc[0] >= 2500 else "양호",
            "취약" if safe_numeric(pd.Series([row.get("fta_ratio", np.nan)])).iloc[0] <= 30 else "주의" if safe_numeric(pd.Series([row.get("fta_ratio", np.nan)])).iloc[0] <= 60 else "양호",
            "취약" if safe_numeric(pd.Series([row.get("지역권수", np.nan)])).iloc[0] <= 3 else "주의" if safe_numeric(pd.Series([row.get("지역권수", np.nan)])).iloc[0] <= 6 else "양호",
            "취약" if safe_numeric(pd.Series([row.get("수입국수", np.nan)])).iloc[0] <= 3 else "주의" if safe_numeric(pd.Series([row.get("수입국수", np.nan)])).iloc[0] <= 6 else "양호",
        ]
    })
    st.dataframe(diagnose_df, use_container_width=True)
    st.write(build_priority_comment(row))

    st.markdown("#### 간이 시뮬레이터")
    st.caption("최종위험점수는 4대 축 점수 변화에 대한 참고용 추정치입니다.")
    w_final = get_entropy_weights(entropy, "PANEL_FINAL")
    if not w_final:
        w_final = {"가격리스크점수": 0.25, "수급리스크점수": 0.25, "물류리스크점수": 0.25, "정책이벤트리스크점수": 0.25}

    price_delta = st.slider("가격리스크점수 변화", -30.0, 30.0, 0.0, 1.0)
    supply_delta = st.slider("수급리스크점수 변화", -30.0, 30.0, 0.0, 1.0)
    logi_delta = st.slider("물류리스크점수 변화", -30.0, 30.0, 0.0, 1.0)
    policy_delta = st.slider("정책이벤트리스크점수 변화", -30.0, 30.0, 0.0, 1.0)

    cur_axes = {
        "가격리스크점수": float(row.get("가격리스크점수", 0) or 0),
        "수급리스크점수": float(row.get("수급리스크점수", 0) or 0),
        "물류리스크점수": float(row.get("물류리스크점수", 0) or 0),
        "정책이벤트리스크점수": float(row.get("정책이벤트리스크점수", 0) or 0),
    }
    new_axes = {
        "가격리스크점수": np.clip(cur_axes["가격리스크점수"] + price_delta, 0, 100),
        "수급리스크점수": np.clip(cur_axes["수급리스크점수"] + supply_delta, 0, 100),
        "물류리스크점수": np.clip(cur_axes["물류리스크점수"] + logi_delta, 0, 100),
        "정책이벤트리스크점수": np.clip(cur_axes["정책이벤트리스크점수"] + policy_delta, 0, 100),
    }

    base_raw_all = (
        safe_numeric(panel["가격리스크점수"]).fillna(0) * w_final.get("가격리스크점수", 0.25) +
        safe_numeric(panel["수급리스크점수"]).fillna(0) * w_final.get("수급리스크점수", 0.25) +
        safe_numeric(panel["물류리스크점수"]).fillna(0) * w_final.get("물류리스크점수", 0.25) +
        safe_numeric(panel["정책이벤트리스크점수"]).fillna(0) * w_final.get("정책이벤트리스크점수", 0.25)
    )
    raw_min = base_raw_all.min()
    raw_max = base_raw_all.max()
    cur_raw = sum(cur_axes[k] * w_final.get(k, 0.25) for k in cur_axes)
    new_raw = sum(new_axes[k] * w_final.get(k, 0.25) for k in new_axes)
    if pd.isna(raw_min) or pd.isna(raw_max) or abs(raw_max - raw_min) < 1e-12:
        cur_final_est = row.get("최종위험점수", np.nan)
        new_final_est = row.get("최종위험점수", np.nan)
    else:
        cur_final_est = (cur_raw - raw_min) / (raw_max - raw_min) * 100
        new_final_est = (new_raw - raw_min) / (raw_max - raw_min) * 100

    sim_df = pd.DataFrame({
        "리스크축": list(cur_axes.keys()),
        "현재값": list(cur_axes.values()),
        "조정후": list(new_axes.values()),
        "가중치": [w_final.get(k, np.nan) for k in cur_axes.keys()]
    })
    st.dataframe(sim_df, use_container_width=True)

    s1, s2 = st.columns(2)
    s1.metric("현재 추정 최종위험점수", fmt_num(cur_final_est, 2))
    s2.metric("조정 후 추정 최종위험점수", fmt_num(new_final_est, 2), delta=fmt_num(new_final_est - cur_final_est, 2))

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
if menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")

    if country is None:
        st.error("COUNTRY_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(country), key="alt_chain")
    month_list = get_month_list(country[country["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=len(month_list)-1 if month_list else 0, key="alt_month")

    sub = country[(country["체인구분"] == chain) & (country["연월"] == month)].copy()
    if sub.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        st.stop()

    fta_only = st.checkbox("FTA 체결국만 보기", value=False)
    exclude_top = st.checkbox("상위공급국 제외", value=False)

    if fta_only and "FTA여부" in sub.columns:
        sub = sub[sub["FTA여부"] == "Y"].copy()
    if exclude_top and "상위공급국여부" in sub.columns:
        sub = sub[sub["상위공급국여부"] != "Y"].copy()

    score = pd.DataFrame(index=sub.index)
    score["risk_rev"] = 100 - safe_numeric(sub["최종보정점수"]).rank(pct=True) * 100
    score["fta_bonus"] = np.where(sub["FTA여부"] == "Y", 20, 0) if "FTA여부" in sub.columns else 0
    score["import_base"] = safe_numeric(sub["국가수입금액"]).rank(pct=True) * 20 if "국가수입금액" in sub.columns else 0
    sub["대체국추천점수"] = score.sum(axis=1)
    sub = sub.sort_values(["대체국추천점수", "최종보정점수"], ascending=[False, True])

    st.markdown("#### 추천 결과")
    show_cols = [c for c in [
        "국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부",
        "국가수입금액", "금액비중", "기본평가점수", "최종보정점수", "최종판정", "대체국추천점수"
    ] if c in sub.columns]
    st.dataframe(sub[show_cols], use_container_width=True)

    top10 = sub.head(10)
    if not top10.empty:
        fig = px.bar(top10, x="국가명", y="대체국추천점수", color="최종판정" if "최종판정" in top10.columns else None)
        fig.update_layout(height=420)
        st.plotly_chart(fig, use_container_width=True)

    show_df_download_button(sub[show_cols], "대체국 추천 결과 CSV 다운로드", f"alternative_country_{chain}_{month}.csv")

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
if menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")

    available_sheet_options = {k: v for k, v in sheet_map.items() if v is not None}
    option_labels = [f"{v} ({k})" for k, v in available_sheet_options.items()]
    reverse_map = {f"{v} ({k})": k for k, v in available_sheet_options.items()}

    selected_label = st.selectbox("시트 선택", option_labels)
    key = reverse_map[selected_label]
    df = data.get(key)

    st.write(f"선택 시트: **{available_sheet_options[key]}**")
    st.write(f"행 수: {len(df):,} / 열 수: {len(df.columns):,}")
    st.dataframe(df.head(200), use_container_width=True)

    show_df_download_button(df, "현재 시트 CSV 다운로드", f"{available_sheet_options[key]}.csv")

    workbook_bytes = to_download_excel({sheet_map[k]: data[k] for k in available_sheet_options.keys() if data.get(k) is not None})
    st.download_button(
        "로드된 전체 시트 Excel로 다시 다운로드",
        data=workbook_bytes,
        file_name="dashboard_loaded_workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# 9. 데이터 검증 / 방법론
# =========================================================
if menu == "9. 데이터 검증 / 방법론":
    st.header("9. 데이터 검증 / 방법론")

    checks = run_checks(sheet_names, sheet_map, data)
    st.markdown("#### 데이터 검증 결과")
    st.dataframe(checks, use_container_width=True)

    if method is not None:
        st.markdown("#### METHOD_GUIDE")
        st.dataframe(method, use_container_width=True)

    if norm_check is not None:
        st.markdown("#### NOMALIZATION_CHECK")
        st.dataframe(norm_check, use_container_width=True)

    if norm_audit is not None:
        st.markdown("#### NOMALIZATION_AUDIT")
        st.dataframe(norm_audit, use_container_width=True)

    if entropy is not None:
        st.markdown("#### ENTROPY_WEIGHT")
        st.dataframe(entropy, use_container_width=True)

    st.markdown("#### 시트 매핑")
    map_df = pd.DataFrame({
        "logical_key": list(sheet_map.keys()),
        "loaded_sheet": list(sheet_map.values())
    })
    st.dataframe(map_df, use_container_width=True)
