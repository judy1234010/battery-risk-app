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

DEFAULT_MONTH = "2025-12"

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
            df["연월"] = df.apply(lambda r: safe_ym(r["연월"], r["연도"], r["월"]), axis=1)
        else:
            df["연월"] = df["연월"].apply(safe_ym)
        return df

    if "연도" in cols and "월" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r["연도"], r["월"]), axis=1)
        return df

    if "연" in cols and "월" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r["연"], r["월"]), axis=1)
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
    try:
        return f"{float(v):.{digits}f}%"
    except Exception:
        return str(v)

def get_default_index(options, default_value=DEFAULT_MONTH):
    if not options:
        return 0
    if default_value in options:
        return options.index(default_value)
    return len(options) - 1

def get_month_list(df):
    if df is None or "연월" not in df.columns:
        return []
    vals = sorted(df["연월"].dropna().astype(str).unique().tolist())
    return vals

def get_chain_list(df):
    if df is None or "체인구분" not in df.columns:
        return []
    return sorted(df["체인구분"].dropna().astype(str).unique().tolist())

def to_download_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            if df is None or not isinstance(df, pd.DataFrame):
                continue
            temp = df.copy()
            temp.to_excel(writer, sheet_name=str(sname)[:31], index=False)
    return output.getvalue()

def show_df_download_button(df, label, filename):
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(label=label, data=csv, file_name=filename, mime="text/csv")

def prepare_aliases(data):
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
            "보정사유", "지역권", "품목명", "품목군", "배터리유형", "비고",
            "선정이유", "해석", "최종해석", "구분", "포함여부", "종합_시사점"
        ]:
            df = normalize_text_col(df, txt_col)

        data[key] = df

    data = prepare_aliases(data)
    return sheet_names, sheet_map, data

def run_checks(sheet_names, sheet_map, data):
    rows = []

    def add(level, item, detail):
        rows.append({"레벨": level, "점검항목": item, "상세": detail})

    required_keys = [
        "country", "panel", "alert", "chain_compare", "entropy",
        "norm_check", "method", "signal_scope", "signal_base",
        "signal_lag", "lead_detail", "lead_compare"
    ]
    for k in required_keys:
        if data.get(k) is None:
            add("FAIL", "필수 시트", f"{k} 시트 누락")
        else:
            add("PASS", "필수 시트", f"{sheet_map[k]} 로딩")

    panel = data.get("panel")
    alert = data.get("alert")
    tpu = data.get("tpu")
    gscpi = data.get("gscpi")
    lead_detail = data.get("lead_detail")
    lead_compare = data.get("lead_compare")
    signal_base = data.get("signal_base")

    if panel is not None:
        panel_need = [
            "연월", "체인구분", "HHI", "상위1국의존도", "경보점수기초", "국가보정합계",
            "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
            "최종위험점수", "최종경보등급", "fta_ratio", "지역권수", "TPU_INDEX_NORM"
        ]
        miss = [c for c in panel_need if c not in panel.columns]
        add("PASS" if not miss else "FAIL", "PANEL_MONTHLY 필수컬럼", "정상" if not miss else f"누락: {miss}")
        add("PASS" if "CV" not in panel.columns else "WARN", "CV 제거 여부", "CV 없음" if "CV" not in panel.columns else "CV 컬럼 잔존")

    if tpu is not None:
        ok = {"TPU_INDEX", "TPU_INDEX_NORM"}.issubset(tpu.columns)
        add("PASS" if ok else "FAIL", "TPU_INDEX 컬럼", "TPU_INDEX / TPU_INDEX_NORM 존재" if ok else "TPU 컬럼 누락")

    if gscpi is not None:
        ok = ("GSCPI" in gscpi.columns) and ("GSCPI_Norm" in gscpi.columns or "GSCPI_NORM" in gscpi.columns)
        add("PASS" if ok else "FAIL", "GSCPI 컬럼", "GSCPI / GSCPI_Norm 존재" if ok else "GSCPI 컬럼 누락")

    if panel is not None and alert is not None:
        compare_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수", "최종경보등급", "fta_ratio"]
        merged = alert.merge(panel[["연월", "체인구분"] + compare_cols], on=["연월", "체인구분"], how="left", suffixes=("_alert", "_panel"))
        mismatch_count = 0
        for c in compare_cols:
            a = f"{c}_alert"
            p = f"{c}_panel"
            if c == "최종경보등급":
                mismatch_count += int((merged[a].fillna("") != merged[p].fillna("")).sum())
            else:
                aa = safe_numeric(merged[a]).round(4)
                pp = safe_numeric(merged[p]).round(4)
                mismatch_count += int((aa.fillna(-9999) != pp.fillna(-9999)).sum())
        add("PASS" if mismatch_count == 0 else "WARN", "ALERT vs PANEL 일치", "핵심 값 일치" if mismatch_count == 0 else f"불일치 셀 {mismatch_count}개")

    if lead_detail is not None:
        if "lag개월" in lead_detail.columns:
            lags = sorted(lead_detail["lag개월"].dropna().astype(int).unique().tolist())
            add("PASS" if lags == [1, 2, 3, 4, 5, 6] else "WARN", "LEAD_SIGNAL_LAG_DETAIL", f"lag 범위: {lags}")
        else:
            add("FAIL", "LEAD_SIGNAL_LAG_DETAIL", "lag개월 컬럼 누락")

    if lead_compare is not None:
        need = ["지표", "최적lag개월", "최적lag상관", "최적lag절대상관", "최종해석"]
        miss = [c for c in need if c not in lead_compare.columns]
        add("PASS" if not miss else "FAIL", "LEAD_SIGNAL_LAG_COMPARE", "정상" if not miss else f"누락: {miss}")

    if signal_base is not None:
        add("PASS" if "최종위험점수_V2" in signal_base.columns else "FAIL", "SIGNAL_BASE 종속변수", "최종위험점수_V2 존재" if "최종위험점수_V2" in signal_base.columns else "최종위험점수_V2 누락")

    return pd.DataFrame(rows)

def get_entropy_weights(entropy_df, stage_name):
    if entropy_df is None or entropy_df.empty or "단계" not in entropy_df.columns:
        return {}
    sub = entropy_df[entropy_df["단계"] == stage_name].copy()
    if sub.empty:
        return {}
    return dict(zip(sub["변수명"], safe_numeric(sub["가중치"])))

def get_priority_zero_reason(alert_df, chain_name):
    if alert_df is None or alert_df.empty:
        return None
    sub = alert_df[alert_df["체인구분"] == chain_name].copy()
    if sub.empty:
        return None

    ratio = (sub["상대적_우선관리대상"] == "Y").mean() * 100 if "상대적_우선관리대상" in sub.columns else np.nan
    if pd.notna(ratio) and ratio > 0:
        return None

    very_high = (sub["최종경보등급"] == "매우높음").sum() if "최종경보등급" in sub.columns else 0
    both = ((sub["최종경보등급"] == "매우높음") & (sub["대체조달가능성"] == "취약")).sum() if {"최종경보등급", "대체조달가능성"}.issubset(sub.columns) else 0

    return (
        f"{chain_name}의 우선관리대상 비중이 0%인 이유는 분석기간 동안 "
        f"'상대적_우선관리대상=Y'로 분류된 월이 없었기 때문입니다. "
        f"매우높음 경보 월은 {very_high}회였고, 동시에 대체조달 취약 조건까지 겹친 사례는 {both}회였습니다."
    )

def build_executive_comment(row):
    comments = []
    final_risk = pd.to_numeric(pd.Series([row.get("최종위험점수", np.nan)]), errors="coerce").iloc[0]
    top1 = pd.to_numeric(pd.Series([row.get("상위1국의존도", np.nan)]), errors="coerce").iloc[0]
    hhi = pd.to_numeric(pd.Series([row.get("HHI", np.nan)]), errors="coerce").iloc[0]
    fta = pd.to_numeric(pd.Series([row.get("fta_ratio", np.nan)]), errors="coerce").iloc[0]
    regions = pd.to_numeric(pd.Series([row.get("지역권수", np.nan)]), errors="coerce").iloc[0]
    alt = pd.to_numeric(pd.Series([row.get("대체조달가능성_점수", np.nan)]), errors="coerce").iloc[0]

    if pd.notna(final_risk) and final_risk >= 60:
        comments.append("현재 최종위험 수준이 높아 단기 대응계획과 월간 모니터링 강화가 필요합니다.")
    if pd.notna(top1) and top1 >= 70:
        comments.append("상위 1개국 의존도가 높아 특정 국가 차질 시 영향이 빠르게 확대될 수 있습니다.")
    if pd.notna(hhi) and hhi >= 5000:
        comments.append("공급국 집중도가 높아 신규 공급선 테스트와 계약 분산 검토가 필요합니다.")
    if pd.notna(fta) and fta <= 30:
        comments.append("FTA 활용비중이 낮아 원가·통관 측면의 개선여지가 큽니다.")
    if pd.notna(regions) and regions <= 3:
        comments.append("지역권 수가 적어 권역 다변화 전략이 필요합니다.")
    if pd.notna(alt) and alt <= 35:
        comments.append("대체조달가능성이 낮아 예비 공급선 확보를 우선과제로 두는 것이 바람직합니다.")
    if not comments:
        comments.append("현재는 즉각적인 구조조정보다 정기 모니터링과 부분 개선이 적절합니다.")
    return " ".join(comments)

def build_action_roadmap(row):
    actions = []
    top1 = pd.to_numeric(pd.Series([row.get("상위1국의존도", np.nan)]), errors="coerce").iloc[0]
    hhi = pd.to_numeric(pd.Series([row.get("HHI", np.nan)]), errors="coerce").iloc[0]
    fta = pd.to_numeric(pd.Series([row.get("fta_ratio", np.nan)]), errors="coerce").iloc[0]
    countries = pd.to_numeric(pd.Series([row.get("수입국수", np.nan)]), errors="coerce").iloc[0]
    regions = pd.to_numeric(pd.Series([row.get("지역권수", np.nan)]), errors="coerce").iloc[0]
    reason = str(row.get("보정사유", ""))

    if pd.notna(top1) and top1 >= 70:
        actions.append(("1순위", "상위국 의존도 완화", "상위 1개국 물량 일부를 2~3위국 또는 신규국으로 분산", "상위1국의존도, HHI 개선"))
    if pd.notna(hhi) and hhi >= 5000:
        actions.append(("1순위", "공급선 집중 완화", "단일 공급국·소수 공급국 구조를 다변화", "HHI, 수입국수 개선"))
    if pd.notna(fta) and fta <= 40:
        actions.append(("2순위", "FTA 활용 재점검", "FTA 활용 가능한 국가/품목의 관세·원산지 조건 검토", "fta_ratio 개선"))
    if pd.notna(regions) and regions <= 3:
        actions.append(("2순위", "권역 다변화", "동일 권역 집중 시 대체 권역 공급선 확보", "지역권수 개선"))
    if pd.notna(countries) and countries <= 4:
        actions.append(("2순위", "예비 공급선 발굴", "소량 테스트 발주를 통해 신규국 실효성 검증", "수입국수 개선"))
    if "정책" in reason:
        actions.append(("모니터링", "정책 이벤트 대응", "통상규제·정책 발표 일정 중심으로 이벤트 캘린더 운영", "정책 충격 조기 대응"))
    if "물류" in reason:
        actions.append(("모니터링", "물류 병목 대응", "운임·리드타임·적체 지표를 월별 추적", "물류 차질 조기 대응"))
    if "가격" in reason:
        actions.append(("모니터링", "가격 헤지 검토", "환율·원자재 가격 변동성에 대한 구매 시점 분산", "가격 충격 완화"))
    if "수급" in reason:
        actions.append(("모니터링", "공급구조 리밸런싱", "고의존 공급선의 계약 구조와 대체조달 가능성을 동시 점검", "수급축 완화"))
    if not actions:
        actions.append(("기본", "정기 모니터링", "월별 리스크 추이와 공급국 변동을 지속 점검", "현 수준 유지 관리"))

    df = pd.DataFrame(actions, columns=["우선순위", "권고 액션", "실행 아이디어", "예상 개선 지표"])
    return df.drop_duplicates()

def apply_scenario_to_row(row_dict, scenario_name):
    row = dict(row_dict)

    def g(k, default=0):
        val = pd.to_numeric(pd.Series([row.get(k, default)]), errors="coerce").iloc[0]
        return default if pd.isna(val) else float(val)

    if scenario_name == "보수적 분산":
        row["수급리스크점수"] = np.clip(g("수급리스크점수") - 8, 0, 100)
        row["상위1국의존도"] = np.clip(g("상위1국의존도") - 8, 0, 100)
        row["HHI"] = max(g("HHI") - 400, 0)
        row["수입국수"] = g("수입국수") + 1
        row["지역권수"] = g("지역권수") + 1
    elif scenario_name == "공격적 분산":
        row["수급리스크점수"] = np.clip(g("수급리스크점수") - 15, 0, 100)
        row["상위1국의존도"] = np.clip(g("상위1국의존도") - 15, 0, 100)
        row["HHI"] = max(g("HHI") - 1000, 0)
        row["수입국수"] = g("수입국수") + 2
        row["지역권수"] = g("지역권수") + 1
        row["fta_ratio"] = np.clip(g("fta_ratio") + 5, 0, 100)
    elif scenario_name == "FTA 확대":
        row["fta_ratio"] = np.clip(g("fta_ratio") + 15, 0, 100)
    elif scenario_name == "물류 완화":
        row["물류리스크점수"] = np.clip(g("물류리스크점수") - 10, 0, 100)
    elif scenario_name == "정책 충격 완화":
        row["정책이벤트리스크점수"] = np.clip(g("정책이벤트리스크점수") - 10, 0, 100)

    return row

def estimate_final_score(panel_df, row_dict, entropy_df=None):
    w = get_entropy_weights(entropy_df, "PANEL_FINAL")
    if not w:
        w = {
            "가격리스크점수": 0.25,
            "수급리스크점수": 0.25,
            "물류리스크점수": 0.25,
            "정책이벤트리스크점수": 0.25,
        }

    base_raw = (
        safe_numeric(panel_df["가격리스크점수"]).fillna(0) * w.get("가격리스크점수", 0.25) +
        safe_numeric(panel_df["수급리스크점수"]).fillna(0) * w.get("수급리스크점수", 0.25) +
        safe_numeric(panel_df["물류리스크점수"]).fillna(0) * w.get("물류리스크점수", 0.25) +
        safe_numeric(panel_df["정책이벤트리스크점수"]).fillna(0) * w.get("정책이벤트리스크점수", 0.25)
    )
    raw_min = base_raw.min()
    raw_max = base_raw.max()

    def g(k):
        val = pd.to_numeric(pd.Series([row_dict.get(k, 0)]), errors="coerce").iloc[0]
        return 0 if pd.isna(val) else float(val)

    new_raw = (
        g("가격리스크점수") * w.get("가격리스크점수", 0.25) +
        g("수급리스크점수") * w.get("수급리스크점수", 0.25) +
        g("물류리스크점수") * w.get("물류리스크점수", 0.25) +
        g("정책이벤트리스크점수") * w.get("정책이벤트리스크점수", 0.25)
    )
    if pd.isna(raw_min) or pd.isna(raw_max) or abs(raw_max - raw_min) < 1e-12:
        return np.nan
    return float(np.clip((new_raw - raw_min) / (raw_max - raw_min) * 100, 0, 100))

# =========================================================
# 메인
# =========================================================
st.title("망보는사람들 공급망 리스크 대시보드")

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
    st.info("이 메뉴는 선택한 기준 연월의 전체 체인 상태를 빠르게 비교하는 화면입니다. 월별 선택이 가능하며, 업로드된 파일의 전체 기간을 모두 탐색할 수 있습니다.")

    if panel is None or chain_compare is None:
        st.error("PANEL_MONTHLY 또는 체인별 비교표 시트가 없습니다.")
        st.stop()

    all_months = get_month_list(panel)
    month = st.selectbox("기준 연월 선택", all_months, index=get_default_index(all_months))
    if all_months:
        st.caption(f"분석 가능 기간: {all_months[0]} ~ {all_months[-1]} · 향후 더 이른/늦은 연월 데이터가 추가돼도 동일 구조로 확장 가능합니다.")

    month_df = panel[panel["연월"] == month].copy()
    if month_df.empty:
        st.warning("선택한 연월의 데이터가 없습니다.")
        st.stop()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기준 연월", month)
    c2.metric("분석 체인 수", month_df["체인구분"].nunique())
    c3.metric("평균 최종위험점수", fmt_num(safe_numeric(month_df["최종위험점수"]).mean(), 2))
    c4.metric("매우높음 체인 수", int((month_df["최종경보등급"] == "매우높음").sum()))

    st.markdown("#### 기준 연월 체인별 요약")
    show_cols = [
        "체인구분", "가격리스크점수", "수급리스크점수", "물류리스크점수",
        "정책이벤트리스크점수", "최종위험점수", "최종경보등급",
        "상위1국의존도", "HHI", "fta_ratio", "지역권수"
    ]
    st.dataframe(month_df[show_cols], use_container_width=True)

    st.markdown("#### 체인별 최종위험 추이")
    fig = px.line(panel.sort_values("연월"), x="연월", y="최종위험점수", color="체인구분", markers=True)
    fig.update_layout(height=420)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 기준 연월 4대 리스크 축 비교")
    axis_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    melt_df = month_df.melt(id_vars=["체인구분"], value_vars=axis_cols, var_name="리스크축", value_name="점수")
    fig2 = px.bar(melt_df, x="리스크축", y="점수", color="체인구분", barmode="group")
    fig2.update_layout(height=420)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 체인별 비교표")
    st.dataframe(chain_compare, use_container_width=True)

    if alert is not None and "우선관리대상_비중" in chain_compare.columns:
        reasons = []
        for _, r in chain_compare.iterrows():
            try:
                val = float(r.get("우선관리대상_비중", 0) or 0)
            except Exception:
                val = np.nan
            if pd.notna(val) and val == 0:
                msg = get_priority_zero_reason(alert, r["체인구분"])
                if msg:
                    reasons.append(msg)
        for msg in reasons:
            st.caption(f"※ {msg}")

    st.markdown("#### 리스크 산식 기준")
    risk_basis_df = pd.DataFrame({
        "리스크축": ["가격리스크", "수급리스크", "물류리스크", "정책이벤트리스크"],
        "주요 입력값": [
            "환율정규화, 납가격정규화, 리튬가격정규화, 니켈가격정규화",
            "HHI·상위1국의존도 기반 경보점수기초 + 국가보정합계",
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
    st.info("이 메뉴는 특정 체인의 장기 추세와 구조적 취약성을 보는 화면입니다. 선택한 연월 수치와 함께 전체 기간의 추이를 동시에 확인할 수 있습니다. 반면 '충격 원인 추적'은 특정 월의 급등 원인을 해석하는 데 집중합니다.")

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chains = get_chain_list(panel)
    chain = st.selectbox("체인 선택", chains)
    sub = panel[panel["체인구분"] == chain].sort_values("연월").copy()
    month_list = get_month_list(sub)
    month = st.selectbox("기준 연월 선택", month_list, index=get_default_index(month_list))
    row_df = sub[sub["연월"] == month]
    if row_df.empty:
        st.warning("선택한 연월 데이터가 없습니다.")
        st.stop()
    row = row_df.iloc[0]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기준 최종위험점수", fmt_num(row["최종위험점수"], 2), row["최종경보등급"])
    c2.metric("상위1국의존도", fmt_pct(row["상위1국의존도"], 2))
    c3.metric("HHI", fmt_num(row["HHI"], 2))
    c4.metric("FTA 비중", fmt_pct(row["fta_ratio"], 2))
    st.caption("위 4개 수치는 선택한 기준 연월 값을 보여줍니다.")

    trend_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수"]
    melt = sub.melt(id_vars=["연월"], value_vars=trend_cols, var_name="지표", value_name="값")
    fig = px.line(melt, x="연월", y="값", color="지표", markers=True)
    fig.update_layout(height=460)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 체인 구조 진단")
    diag = pd.DataFrame({
        "항목": ["월평균 수입금액", "평균 상위1국의존도", "평균 HHI", "평균 FTA 활용비중", "평균 지역권수"],
        "값": [
            fmt_num(safe_numeric(sub["총수입금액"]).mean(), 2),
            fmt_pct(safe_numeric(sub["상위1국의존도"]).mean(), 2),
            fmt_num(safe_numeric(sub["HHI"]).mean(), 2),
            fmt_pct(safe_numeric(sub["fta_ratio"]).mean(), 2),
            fmt_num(safe_numeric(sub["지역권수"]).mean(), 2),
        ]
    })
    st.dataframe(diag, use_container_width=True)

    st.markdown("#### 월별 상세 데이터")
    display_cols = [
        "연월", "총수입금액", "총수입물량", "평균수입단가", "수입국수", "지역권수",
        "상위1국의존도", "상위3국집중도", "HHI", "경보점수기초", "국가보정합계",
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
        "최종위험점수", "최종경보등급"
    ]
    st.dataframe(sub[display_cols], use_container_width=True)

    if hs_summary is not None and "체인구분" in hs_summary.columns:
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
    st.info("이 메뉴는 특정 체인·연월에서 어떤 국가가 실제 공급구조를 좌우하는지 확인하는 화면입니다. 국가별 위험, 비중, FTA 여부, 집중보정 효과를 함께 봅니다.")

    if country is None:
        st.error("COUNTRY_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(country), key="country_chain")
    month_list = get_month_list(country[country["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=get_default_index(month_list))

    sub = country[(country["체인구분"] == chain) & (country["연월"] == month)].copy()
    sub = sub.sort_values(["금액비중", "국가수입금액"], ascending=[False, False])

    c1, c2, c3 = st.columns(3)
    c1.metric("공급국 수", sub["국가명"].nunique())
    c2.metric("평균 최종보정점수", fmt_num(safe_numeric(sub["최종보정점수"]).mean(), 2))
    c3.metric("상위공급국 수", int((sub["상위공급국여부"] == "Y").sum()) if "상위공급국여부" in sub.columns else 0)

    st.caption("최종보정점수는 국가기본위험에 공급집중·지역집중·상위공급국 여부·FTA 위험 등의 보정요인을 반영한 최종 국가위험 점수입니다. 값이 높을수록 해당 국가를 통한 조달 리스크가 더 크다는 의미입니다.")

    topn = st.slider("상위 몇 개 국가를 볼지", 5, min(30, len(sub)) if len(sub) > 0 else 5, 10)
    top = sub.head(topn)

    fig = px.bar(
        top,
        x="국가명",
        y="금액비중",
        color="최종판정" if "최종판정" in top.columns else None,
        hover_data=[c for c in ["국가코드", "FTA여부", "최종보정점수"] if c in top.columns]
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
    st.info("이 메뉴는 특정 체인·특정 월의 위험 급등이 어느 축에서 발생했는지 해석하는 화면입니다. '체인별 심층 분석'이 장기 추세 중심이라면, 여기서는 선택 월의 직접 원인을 읽는 데 초점을 둡니다.")

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel), key="shock_chain")
    month_list = get_month_list(panel[panel["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=get_default_index(month_list), key="shock_month")

    row_df = panel[(panel["체인구분"] == chain) & (panel["연월"] == month)].copy()
    if row_df.empty:
        st.warning("선택 조건에 해당하는 데이터가 없습니다.")
        st.stop()
    row = row_df.iloc[0]

    axis_df = pd.DataFrame({
        "리스크축": ["가격리스크", "수급리스크", "물류리스크", "정책이벤트리스크"],
        "점수": [
            row.get("가격리스크점수", np.nan),
            row.get("수급리스크점수", np.nan),
            row.get("물류리스크점수", np.nan),
            row.get("정책이벤트리스크점수", np.nan),
        ]
    }).sort_values("점수", ascending=False)

    if alert is not None:
        arow = alert[(alert["체인구분"] == chain) & (alert["연월"] == month)]
        if not arow.empty:
            merged = {**row.to_dict(), **arow.iloc[0].to_dict()}
            row = pd.Series(merged)

    c1, c2, c3 = st.columns(3)
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan), 2), row.get("최종경보등급", "-"))
    c2.metric("주요 원인축", axis_df.iloc[0]["리스크축"])
    c3.metric("보정사유", row.get("보정사유", "-"))

    fig = px.bar(axis_df, x="리스크축", y="점수", color="리스크축")
    fig.update_layout(height=380, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 선택 월 원인 해석")
    reasons = [
        f"- 가격리스크점수: {fmt_num(row.get('가격리스크점수', np.nan), 2)}",
        f"- 수급리스크점수: {fmt_num(row.get('수급리스크점수', np.nan), 2)} (HHI {fmt_num(row.get('HHI', np.nan), 2)}, 상위1국의존도 {fmt_pct(row.get('상위1국의존도', np.nan), 2)}, 국가보정합계 {fmt_num(row.get('국가보정합계', np.nan), 2)})",
        f"- 물류리스크점수: {fmt_num(row.get('물류리스크점수', np.nan), 2)} (GSCPI_Norm 기준)",
        f"- 정책이벤트리스크점수: {fmt_num(row.get('정책이벤트리스크점수', np.nan), 2)} (TPU_INDEX_NORM 기준)"
    ]
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
            st.info(f"TPU 원천값: {fmt_num(t.get('TPU_INDEX', np.nan), 2)} / 정규화값(TPU_INDEX_NORM): {fmt_num(t.get('TPU_INDEX_NORM', np.nan), 2)}")
            if "서사배경" in t.index:
                st.markdown("#### 정책 이벤트 배경")
                st.write(t.get("서사배경", "-"))

# =========================================================
# 5. 선행 신호 후보 탐지
# =========================================================
if menu == "5. 선행 신호 후보 탐지":
    st.header("5. 선행 신호 후보 탐지")
    st.info("이 메뉴는 원재료·원천 변수가 몇 개월 뒤 최종위험과 가장 강하게 연결되는지를 lag 1~6 기준으로 비교하는 화면입니다. 실무적으로는 '무엇을 먼저 모니터링할지' 우선순위를 정할 때 활용합니다.")

    if signal_scope is None or lead_compare is None or lead_detail is None:
        st.error("SIGNAL / LEAD_SIGNAL 관련 시트가 없습니다.")
        st.stop()

    chain_options = sorted(lead_compare["체인구분"].dropna().astype(str).unique().tolist())
    chain = st.selectbox("체인 선택", chain_options, key="lead_chain")

    st.markdown("#### 1) 후보 변수 범위")
    st.caption("어떤 변수들이 선행신호 후보로 포함되었는지 확인하는 표입니다. 포함여부(Y/N), 구분, 비고를 통해 변수의 의미와 사용 범위를 빠르게 파악할 수 있습니다.")
    scope_sub = signal_scope[(signal_scope["체인구분"].isin(["공통", chain])) & (signal_scope["포함여부"].isin(["Y", "N"]))].copy()
    st.dataframe(scope_sub, use_container_width=True)

    st.markdown("#### 2) 변수별 최적 lag 요약")
    st.caption("각 변수에 대해 lag1~lag6 중 절대상관이 가장 큰 시점을 요약한 표입니다. 실무에서는 이 표를 먼저 보고 우선 모니터링 후보를 고르는 것이 가장 효율적입니다.")
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
    st.caption("선택한 지표가 1개월 선행인지, 3개월 선행인지 등 세부 패턴을 확인하는 구간입니다. 상관 방향과 크기를 함께 읽어야 합니다.")
    metric = st.selectbox("지표 선택", compare_sub["지표"].tolist(), key="lead_metric")
    detail_sub = lead_detail[(lead_detail["체인구분"] == chain) & (lead_detail["지표"] == metric)].copy()
    detail_sub = detail_sub.sort_values("lag개월")
    st.dataframe(detail_sub, use_container_width=True)

    fig2 = px.line(detail_sub, x="lag개월", y="상관계수", markers=True)
    fig2.update_layout(height=380, xaxis=dict(dtick=1))
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 4) 검증 베이스 데이터 미리보기")
    st.caption("실제 lag 분석에 사용된 베이스 시계열을 확인하는 구간입니다. 지표 원값과 최종위험점수_V2를 함께 보며 데이터 연결이 자연스러운지 점검할 수 있습니다.")
    if signal_base is not None:
        base_cols = ["연월", "체인구분", "최종위험점수_V2", metric]
        base_cols = [c for c in base_cols if c in signal_base.columns]
        base_sub = signal_base[signal_base["체인구분"] == chain][base_cols].copy()
        st.dataframe(base_sub, use_container_width=True)

# =========================================================
# 6. 기업 대응 우선순위 추천 / 시뮬레이터
# =========================================================
if menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.header("6. 기업 대응 우선순위 추천 / 시뮬레이터")
    st.info("이 메뉴는 기업 실무 관점에서 '지금 무엇을 먼저 바꿔야 하는가'를 정리해주는 실행형 화면입니다. 단순 점수 조회보다, 개선 우선순위·실행 아이디어·시나리오별 효과를 함께 보도록 구성했습니다.")

    if panel is None or alert is None:
        st.error("PANEL_MONTHLY 또는 ALERT_RESULT 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel), key="prio_chain")
    month_list = get_month_list(panel[panel["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=get_default_index(month_list), key="prio_month")

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

    st.markdown("#### 경영진 요약 코멘트")
    st.write(build_executive_comment(row))

    st.markdown("#### 구조 진단")
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
            "취약" if pd.to_numeric(pd.Series([row.get("상위1국의존도", np.nan)]), errors="coerce").iloc[0] >= 70 else "주의" if pd.to_numeric(pd.Series([row.get("상위1국의존도", np.nan)]), errors="coerce").iloc[0] >= 50 else "양호",
            "취약" if pd.to_numeric(pd.Series([row.get("HHI", np.nan)]), errors="coerce").iloc[0] >= 5000 else "주의" if pd.to_numeric(pd.Series([row.get("HHI", np.nan)]), errors="coerce").iloc[0] >= 2500 else "양호",
            "취약" if pd.to_numeric(pd.Series([row.get("fta_ratio", np.nan)]), errors="coerce").iloc[0] <= 30 else "주의" if pd.to_numeric(pd.Series([row.get("fta_ratio", np.nan)]), errors="coerce").iloc[0] <= 60 else "양호",
            "취약" if pd.to_numeric(pd.Series([row.get("지역권수", np.nan)]), errors="coerce").iloc[0] <= 3 else "주의" if pd.to_numeric(pd.Series([row.get("지역권수", np.nan)]), errors="coerce").iloc[0] <= 6 else "양호",
            "취약" if pd.to_numeric(pd.Series([row.get("수입국수", np.nan)]), errors="coerce").iloc[0] <= 3 else "주의" if pd.to_numeric(pd.Series([row.get("수입국수", np.nan)]), errors="coerce").iloc[0] <= 6 else "양호",
        ]
    })
    st.dataframe(diagnose_df, use_container_width=True)

    st.markdown("#### 권고 액션 로드맵")
    roadmap = build_action_roadmap(row)
    st.dataframe(roadmap, use_container_width=True)

    st.markdown("#### 시나리오형 시뮬레이터")
    st.caption("실제 의사결정에서는 '공급선 분산', 'FTA 확대', '물류 완화' 같은 실행 시나리오를 먼저 검토하는 경우가 많습니다. 아래는 대표 시나리오의 참고용 추정 결과입니다.")

    scenario = st.selectbox("대표 시나리오 선택", ["현 상태 유지", "보수적 분산", "공격적 분산", "FTA 확대", "물류 완화", "정책 충격 완화"])
    scenario_row = dict(row)
    if scenario != "현 상태 유지":
        scenario_row = apply_scenario_to_row(row, scenario)

    current_est = estimate_final_score(panel, row, entropy)
    scenario_est = estimate_final_score(panel, scenario_row, entropy)

    sim_df = pd.DataFrame({
        "항목": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "상위1국의존도", "HHI", "FTA 비중", "수입국수", "지역권수"],
        "현재": [
            row.get("가격리스크점수", np.nan),
            row.get("수급리스크점수", np.nan),
            row.get("물류리스크점수", np.nan),
            row.get("정책이벤트리스크점수", np.nan),
            row.get("상위1국의존도", np.nan),
            row.get("HHI", np.nan),
            row.get("fta_ratio", np.nan),
            row.get("수입국수", np.nan),
            row.get("지역권수", np.nan),
        ],
        "시나리오 적용 후": [
            scenario_row.get("가격리스크점수", np.nan),
            scenario_row.get("수급리스크점수", np.nan),
            scenario_row.get("물류리스크점수", np.nan),
            scenario_row.get("정책이벤트리스크점수", np.nan),
            scenario_row.get("상위1국의존도", np.nan),
            scenario_row.get("HHI", np.nan),
            scenario_row.get("fta_ratio", np.nan),
            scenario_row.get("수입국수", np.nan),
            scenario_row.get("지역권수", np.nan),
        ]
    })
    st.dataframe(sim_df, use_container_width=True)

    s1, s2 = st.columns(2)
    s1.metric("현재 추정 최종위험점수", fmt_num(current_est, 2))
    s2.metric("시나리오 적용 후 추정 최종위험점수", fmt_num(scenario_est, 2), delta=fmt_num(scenario_est - current_est, 2) if pd.notna(current_est) and pd.notna(scenario_est) else "-")

    st.markdown("#### 직접 조정 시뮬레이터")
    st.caption("세부 항목을 직접 조정해보면서 어느 변수 개선이 더 효과적인지 탐색할 수 있습니다. 일반적으로 수급리스크 완화, 상위1국의존도 하락, FTA 비중 상승이 구조 개선에 가장 직접적입니다.")

    price_delta = st.slider("가격리스크점수 변화", -30.0, 30.0, 0.0, 1.0)
    supply_delta = st.slider("수급리스크점수 변화", -30.0, 30.0, 0.0, 1.0)
    logi_delta = st.slider("물류리스크점수 변화", -30.0, 30.0, 0.0, 1.0)
    policy_delta = st.slider("정책이벤트리스크점수 변화", -30.0, 30.0, 0.0, 1.0)

    manual_row = dict(row)
    for k, d in {
        "가격리스크점수": price_delta,
        "수급리스크점수": supply_delta,
        "물류리스크점수": logi_delta,
        "정책이벤트리스크점수": policy_delta,
    }.items():
        base_val = pd.to_numeric(pd.Series([manual_row.get(k, 0)]), errors="coerce").iloc[0]
        base_val = 0 if pd.isna(base_val) else float(base_val)
        manual_row[k] = float(np.clip(base_val + d, 0, 100))

    manual_est = estimate_final_score(panel, manual_row, entropy)

    s3, s4 = st.columns(2)
    s3.metric("직접 조정 후 추정 최종위험점수", fmt_num(manual_est, 2))
    s4.metric("현재 대비 변화", fmt_num(manual_est - current_est, 2) if pd.notna(manual_est) and pd.notna(current_est) else "-")

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
if menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")
    st.info("이 메뉴는 현재 공급국 대신 검토할 수 있는 대체 후보를 정리하는 화면입니다. 상위공급국 제외 옵션은 '이미 많이 쓰고 있는 국가'를 빼고 신규 보완 후보를 찾고 싶을 때 활용합니다.")

    if country is None:
        st.error("COUNTRY_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(country), key="alt_chain")
    month_list = get_month_list(country[country["체인구분"] == chain])
    month = st.selectbox("연월 선택", month_list, index=get_default_index(month_list), key="alt_month")

    sub = country[(country["체인구분"] == chain) & (country["연월"] == month)].copy()
    if sub.empty:
        st.warning("조건에 맞는 데이터가 없습니다.")
        st.stop()

    fta_only = st.checkbox("FTA 체결국만 보기", value=False)
    exclude_top = st.checkbox("상위공급국 제외", value=False, help="기존 주력 공급국이 아니라 신규 대체 후보를 찾고 싶을 때 사용합니다.")

    if fta_only and "FTA여부" in sub.columns:
        sub = sub[sub["FTA여부"] == "Y"].copy()
    if exclude_top and "상위공급국여부" in sub.columns:
        sub = sub[sub["상위공급국여부"] != "Y"].copy()

    risk_rank = safe_numeric(sub["최종보정점수"]).rank(pct=True, ascending=True)
    import_rank = safe_numeric(sub["국가수입금액"]).rank(pct=True) if "국가수입금액" in sub.columns else pd.Series(0, index=sub.index)

    sub["대체국추천점수"] = (100 - risk_rank * 100) + np.where(sub["FTA여부"] == "Y", 20, 0) + import_rank * 20
    sub = sub.sort_values(["대체국추천점수", "최종보정점수"], ascending=[False, True])

    st.markdown("#### 추천 결과")
    show_cols = [c for c in [
        "국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부",
        "국가수입금액", "금액비중", "기본평가점수", "최종보정점수", "최종판정", "대체국추천점수"
    ] if c in sub.columns]
    st.dataframe(sub[show_cols], use_container_width=True)

    top10 = sub.head(10)
    if not top10.empty:
        fig = px.bar(top10, x="국가명", y="대체국추천점수", color_discrete_sequence=["#4C78A8"])
        fig.update_layout(height=420, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.caption("막대 색상은 단일 점수(대체국추천점수)만 표현하도록 통일했습니다.")
    show_df_download_button(sub[show_cols], "대체국 추천 결과 CSV 다운로드", f"alternative_country_{chain}_{month}.csv")

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
if menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")
    st.info("대시보드에 연결된 실제 원천 시트를 직접 확인하고 내려받는 메뉴입니다. 조원 검토, 보고서 작성, 후속 분석 연결 시 활용할 수 있습니다.")

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
    st.info("시트 연결 상태, 핵심 컬럼 존재 여부, PANEL/ALERT 정합성, lag 결과 구조 등을 점검하는 메뉴입니다. 앱 배포 전 최종 검토용으로 활용할 수 있습니다.")

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
    st.info("시트 연결 상태, 핵심 컬럼 존재 여부, PANEL/ALERT 정합성
