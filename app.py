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
FTA_COLOR_MAP = {"Y": "#2E8B57", "N": "#4C78A8"}
RISK_COLOR_MAP = {"낮음": "#4C78A8", "보통": "#F2C14E", "높음": "#F28E2B", "매우높음": "#E15759"}

# =========================================================
# 공통 유틸
# =========================================================
def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def safe_numeric(s):
    return pd.to_numeric(s, errors="coerce")

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
            if len(parts) >= 2:
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

def fmt_int(v):
    if pd.isna(v):
        return "-"
    try:
        return f"{int(round(float(v))):,}"
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
    return sorted(df["연월"].dropna().astype(str).unique().tolist())

def get_chain_list(df):
    if df is None or "체인구분" not in df.columns:
        return []
    return sorted(df["체인구분"].dropna().astype(str).unique().tolist())

def grade_final_risk(v):
    v = safe_numeric(pd.Series([v])).iloc[0]
    if pd.isna(v):
        return "해석유보"
    if v >= 80:
        return "매우높음"
    if v >= 60:
        return "높음"
    if v >= 40:
        return "보통"
    return "낮음"

def make_download_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=str(sname)[:31], index=False)
    return output.getvalue()

def show_df_download_button(df, label, filename):
    st.download_button(
        label=label,
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=filename,
        mime="text/csv"
    )

def scenario_rule_text(name):
    mapping = {
        "현 상태 유지": [
            "현재 관측된 공급구조와 리스크 점수를 기준값으로 유지한다."
        ],
        "보수적 분산": [
            "객관 기준: 현재 수입국수에 1개 공급국, 지역권수에 1개 권역을 추가하는 수준의 제한적 다변화를 가정한다.",
            "반영 규칙: 상위1국의존도는 현재 값의 10% 또는 8%p 중 작은 폭만큼 완화하고, HHI는 현재 값의 8% 또는 600 중 작은 폭만큼 완화한다.",
            "의미: 기존 주력 공급선을 유지하되 일부 물량만 보조 공급선으로 분산하는 보수적 전략이다."
        ],
        "공격적 분산": [
            "객관 기준: 현재 수입국수에 2개 공급국, 지역권수에 1개 권역을 추가하는 적극적 다변화를 가정한다.",
            "반영 규칙: 상위1국의존도는 현재 값의 20% 또는 15%p 중 작은 폭만큼 완화하고, HHI는 현재 값의 18% 또는 1,200 중 작은 폭만큼 완화한다.",
            "의미: 신규 공급국과 대체 권역 확보를 통해 공급집중도를 적극적으로 낮추는 전략이다."
        ],
        "FTA 확대": [
            "객관 기준: FTA 미활용 또는 저활용 공급선의 일부 전환을 통해 FTA 활용비중을 현재 대비 10%p 확대하는 경우를 가정한다.",
            "의미: 통관·관세 측면의 완충력을 높이는 전략이다."
        ],
        "물류 완화": [
            "객관 기준: 물류 병목, 적체, 운송 불확실성이 완화되어 물류리스크점수가 10점 감소하는 경우를 가정한다.",
            "의미: 운송 차질 완화, 리드타임 안정화, 공급망 혼잡 해소를 상정한 시나리오이다."
        ],
        "정책 충격 완화": [
            "객관 기준: 정책 이벤트 노출도가 낮은 공급선 비중이 확대되어 정책이벤트리스크점수가 10점 감소하는 경우를 가정한다.",
            "의미: 통상규제·정책 불확실성 노출을 일부 줄이는 구조 조정을 상정한 시나리오이다."
        ],
    }
    return mapping.get(name, [])

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
            "보정사유", "지역권", "품목명", "비고", "해석", "최종해석", "포함여부"
        ]:
            df = normalize_text_col(df, txt_col)
        data[key] = df
    return sheet_names, sheet_map, prepare_aliases(data)

def run_checks(sheet_map, data):
    rows = []
    def add(level, item, detail):
        rows.append({"레벨": level, "점검항목": item, "상세": detail})

    for k in ["country", "panel", "alert", "chain_compare"]:
        if data.get(k) is None:
            add("FAIL", "필수 시트", f"{k} 시트 누락")
        else:
            add("PASS", "필수 시트", f"{sheet_map[k]} 로딩")

    panel = data.get("panel")
    alert = data.get("alert")
    tpu = data.get("tpu")
    lead_detail = data.get("lead_detail")
    lead_compare = data.get("lead_compare")
    signal_base = data.get("signal_base")

    if panel is not None:
        need = ["연월", "체인구분", "최종위험점수", "최종경보등급"]
        miss = [c for c in need if c not in panel.columns]
        add("PASS" if not miss else "FAIL", "PANEL 필수 컬럼", "정상" if not miss else f"누락: {miss}")
        add("PASS" if "CV" not in panel.columns else "WARN", "CV 제거 여부", "CV 없음" if "CV" not in panel.columns else "CV 컬럼 잔존")

    if tpu is not None:
        ok = {"TPU_INDEX", "TPU_INDEX_NORM"}.issubset(tpu.columns)
        add("PASS" if ok else "WARN", "TPU 컬럼", "TPU_INDEX / TPU_INDEX_NORM 확인" if ok else f"현재 컬럼: {list(tpu.columns)}")

    if panel is not None and alert is not None and {"연월", "체인구분"}.issubset(panel.columns) and {"연월", "체인구분"}.issubset(alert.columns):
        cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수", "최종경보등급"] if c in panel.columns and c in alert.columns]
        m = alert[["연월", "체인구분"] + cols].merge(panel[["연월", "체인구분"] + cols], on=["연월", "체인구분"], suffixes=("_alert", "_panel"))
        mismatch = 0
        for c in cols:
            if c == "최종경보등급":
                mismatch += int((m[f"{c}_alert"].fillna("") != m[f"{c}_panel"].fillna("")).sum())
            else:
                a = safe_numeric(m[f"{c}_alert"]).round(4)
                p = safe_numeric(m[f"{c}_panel"]).round(4)
                mismatch += int((a.fillna(-9999) != p.fillna(-9999)).sum())
        add("PASS" if mismatch == 0 else "WARN", "ALERT vs PANEL", "핵심 값 일치" if mismatch == 0 else f"불일치 셀 {mismatch}개")

    if lead_detail is not None and "lag개월" in lead_detail.columns:
        lags = sorted(lead_detail["lag개월"].dropna().astype(int).unique().tolist())
        add("PASS" if lags == [1, 2, 3, 4, 5, 6] else "WARN", "lag 범위", f"{lags}")

    if lead_compare is not None:
        miss = [c for c in ["지표", "최적lag개월", "최적lag상관"] if c not in lead_compare.columns]
        add("PASS" if not miss else "WARN", "LEAD_COMPARE 구조", "정상" if not miss else f"누락: {miss}")

    if signal_base is not None:
        add("PASS" if "최종위험점수_V2" in signal_base.columns else "WARN", "SIGNAL_BASE 종속변수", "최종위험점수_V2 존재" if "최종위험점수_V2" in signal_base.columns else "최종위험점수_V2 누락")

    return pd.DataFrame(rows)

def get_priority_zero_reason(alert_df, chain_name):
    if alert_df is None or alert_df.empty or "체인구분" not in alert_df.columns:
        return None
    sub = alert_df[alert_df["체인구분"] == chain_name].copy()
    if sub.empty or "상대적_우선관리대상" not in sub.columns:
        return None
    ratio = (sub["상대적_우선관리대상"] == "Y").mean() * 100
    if ratio > 0:
        return None
    very_high = (sub["최종경보등급"] == "매우높음").sum() if "최종경보등급" in sub.columns else 0
    return (
        f"{chain_name}의 우선관리대상 비중 0%는 위험 부재가 아니라 우선관리 판정 기준을 충족한 사례가 없었다는 의미이다. "
        f"본 앱에서 우선관리대상은 최종경보등급이 '매우높음'이고 동시에 대체조달가능성이 '취약'인 경우로 정의한다. "
        f"따라서 해당 체인은 매우높음 경보가 {very_high}회 발생했더라도 두 조건이 동시에 성립하지 않아 우선관리대상 비중이 0%로 집계되었다."
    )

def build_executive_comment(row):
    msgs = []
    top1 = safe_numeric(pd.Series([row.get("상위1국의존도", np.nan)])).iloc[0]
    hhi = safe_numeric(pd.Series([row.get("HHI", np.nan)])).iloc[0]
    fta = safe_numeric(pd.Series([row.get("fta_ratio", np.nan)])).iloc[0]
    final = safe_numeric(pd.Series([row.get("최종위험점수", np.nan)])).iloc[0]
    if pd.notna(final) and final >= 60:
        msgs.append("현재 최종위험 수준이 높아 단기 대응계획과 월간 모니터링 강화가 필요하다.")
    if pd.notna(top1) and top1 >= 70:
        msgs.append("상위 1개국 의존도가 높아 특정 국가 차질 시 영향이 확대될 가능성이 크다.")
    if pd.notna(hhi) and hhi >= 5000:
        msgs.append("공급국 집중도가 높아 신규 공급선 테스트와 계약 분산 검토가 요구된다.")
    if pd.notna(fta) and fta <= 30:
        msgs.append("FTA 활용비중이 낮아 비용·통관 측면의 완충여지가 제한적이다.")
    if not msgs:
        msgs.append("현재는 즉각적인 구조조정보다 정기 모니터링과 부분 개선이 적절하다.")
    return " ".join(msgs)

def build_action_roadmap(row):
    items = []
    top1 = safe_numeric(pd.Series([row.get("상위1국의존도", np.nan)])).iloc[0]
    hhi = safe_numeric(pd.Series([row.get("HHI", np.nan)])).iloc[0]
    fta = safe_numeric(pd.Series([row.get("fta_ratio", np.nan)])).iloc[0]
    regions = safe_numeric(pd.Series([row.get("지역권수", np.nan)])).iloc[0]
    if pd.notna(top1) and top1 >= 70:
        items.append(("1순위", "상위국 의존도 완화", "상위 1개국 물량 일부를 2~3위국 또는 신규국으로 분산", "상위1국의존도, HHI 개선"))
    if pd.notna(hhi) and hhi >= 5000:
        items.append(("1순위", "공급선 집중 완화", "단일·소수 공급국 구조를 다변화", "HHI, 수입국수 개선"))
    if pd.notna(fta) and fta <= 40:
        items.append(("2순위", "FTA 활용 재점검", "관세·원산지 조건을 검토해 FTA 활용 확대", "fta_ratio 개선"))
    if pd.notna(regions) and regions <= 3:
        items.append(("2순위", "권역 다변화", "동일 권역 집중 시 대체 권역 공급선 확보", "지역권수 개선"))
    if not items:
        items.append(("기본", "정기 모니터링", "월별 리스크 추이와 공급국 변동을 점검", "현 수준 유지 관리"))
    return pd.DataFrame(items, columns=["우선순위", "권고 액션", "실행 아이디어", "예상 개선 지표"])

def get_entropy_weights(entropy_df):
    if entropy_df is None or entropy_df.empty or "변수명" not in entropy_df.columns or "가중치" not in entropy_df.columns:
        return {}
    return dict(zip(entropy_df["변수명"], safe_numeric(entropy_df["가중치"])))

def estimate_final_score(panel_df, row_dict, entropy_df=None):
    w = get_entropy_weights(entropy_df)
    weights = {
        "가격리스크점수": w.get("가격리스크점수", 0.25),
        "수급리스크점수": w.get("수급리스크점수", 0.25),
        "물류리스크점수": w.get("물류리스크점수", 0.25),
        "정책이벤트리스크점수": w.get("정책이벤트리스크점수", 0.25),
    }
    base_raw = (
        safe_numeric(panel_df.get("가격리스크점수", pd.Series(dtype=float))).fillna(0) * weights["가격리스크점수"] +
        safe_numeric(panel_df.get("수급리스크점수", pd.Series(dtype=float))).fillna(0) * weights["수급리스크점수"] +
        safe_numeric(panel_df.get("물류리스크점수", pd.Series(dtype=float))).fillna(0) * weights["물류리스크점수"] +
        safe_numeric(panel_df.get("정책이벤트리스크점수", pd.Series(dtype=float))).fillna(0) * weights["정책이벤트리스크점수"]
    )
    raw_min, raw_max = base_raw.min(), base_raw.max()

    def g(k):
        v = safe_numeric(pd.Series([row_dict.get(k, 0)])).iloc[0]
        return 0 if pd.isna(v) else float(v)

    new_raw = sum(g(k) * weights[k] for k in weights)
    if pd.isna(raw_min) or pd.isna(raw_max) or abs(raw_max - raw_min) < 1e-12:
        return np.nan
    return float(np.clip((new_raw - raw_min) / (raw_max - raw_min) * 100, 0, 100))

def apply_scenario_to_row(row_dict, scenario_name):
    row = dict(row_dict)

    def g(k):
        v = safe_numeric(pd.Series([row.get(k, 0)])).iloc[0]
        return 0 if pd.isna(v) else float(v)

    top1 = g("상위1국의존도")
    hhi = g("HHI")
    countries = g("수입국수")
    regions = g("지역권수")
    fta = g("fta_ratio")

    if scenario_name == "보수적 분산":
        top1_delta = min(top1 * 0.10, 8.0)
        hhi_delta = min(hhi * 0.08, 600.0)
        row["수급리스크점수"] = np.clip(g("수급리스크점수") - 6, 0, 100)
        row["상위1국의존도"] = np.clip(top1 - top1_delta, 0, 100)
        row["HHI"] = max(hhi - hhi_delta, 0)
        row["수입국수"] = countries + 1
        row["지역권수"] = regions + 1

    elif scenario_name == "공격적 분산":
        top1_delta = min(top1 * 0.20, 15.0)
        hhi_delta = min(hhi * 0.18, 1200.0)
        row["수급리스크점수"] = np.clip(g("수급리스크점수") - 12, 0, 100)
        row["상위1국의존도"] = np.clip(top1 - top1_delta, 0, 100)
        row["HHI"] = max(hhi - hhi_delta, 0)
        row["수입국수"] = countries + 2
        row["지역권수"] = regions + 1
        row["fta_ratio"] = np.clip(fta + 5, 0, 100)

    elif scenario_name == "FTA 확대":
        row["fta_ratio"] = np.clip(fta + 10, 0, 100)

    elif scenario_name == "물류 완화":
        row["물류리스크점수"] = np.clip(g("물류리스크점수") - 10, 0, 100)

    elif scenario_name == "정책 충격 완화":
        row["정책이벤트리스크점수"] = np.clip(g("정책이벤트리스크점수") - 10, 0, 100)

    return row

# =========================================================
# 앱 로드
# =========================================================
st.title("망보는사람들 공급망 리스크 대시보드")

uploaded_file = st.sidebar.file_uploader("최종 엑셀 파일 업로드", type=["xlsx"])
if uploaded_file is None:
    st.info("최종 엑셀 파일(.xlsx)을 업로드하면 대시보드가 열립니다.")
    st.stop()

sheet_names, sheet_map, data = load_workbook(uploaded_file)
check_df = run_checks(sheet_map, data)

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
    st.markdown(
        """
**분석 목적**  
선택한 기준 연월에서 배터리 공급망의 전반적 위험 수준을 체인별로 비교하기 위한 화면이다.

**핵심 구성**  
1. 체인별 최종위험점수 및 경보등급 비교  
2. 4대 리스크 축의 상대적 수준 비교  
3. 체인별 비교표를 통한 평균 위험도와 우선관리 현황 요약

**해석 기준**  
전체 체인 중 상대적으로 위험 수준이 높은 대상을 먼저 식별한 뒤, 이후 심층 분석이 필요한 체인을 선별하는 출발점으로 사용한다.
"""
    )

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    months = get_month_list(panel)
    month = st.selectbox("기준 연월 선택", months, index=get_default_index(months))
    if months:
        st.caption(f"분석 가능 기간: {months[0]} ~ {months[-1]}")

    month_df = panel[panel["연월"] == month].copy()
    if month_df.empty:
        st.warning("선택한 연월 데이터가 없습니다.")
        st.stop()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기준 연월", month)
    c2.metric("분석 체인 수", month_df["체인구분"].nunique())
    c3.metric("평균 최종위험점수", fmt_num(safe_numeric(month_df["최종위험점수"]).mean()))
    c4.metric("매우높음 체인 수", int((month_df["최종경보등급"] == "매우높음").sum()) if "최종경보등급" in month_df.columns else 0)

    show_cols = [c for c in ["체인구분", "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수", "최종경보등급", "상위1국의존도", "HHI", "fta_ratio", "지역권수"] if c in month_df.columns]
    st.markdown("#### 기준 연월 체인별 요약")
    st.dataframe(month_df[show_cols], use_container_width=True)

    st.markdown("#### 4대 리스크 축의 구성 기준")
    risk_basis_df = pd.DataFrame({
        "리스크 축": ["가격리스크", "수급리스크", "물류리스크", "정책이벤트리스크"],
        "구성 기준": [
            "환율정규화, 납가격정규화, 리튬가격정규화, 니켈가격정규화",
            "상위1국의존도, HHI, 경보점수기초, 국가보정합계",
            "GSCPI 및 GSCPI_Norm",
            "TPU_INDEX 및 TPU_INDEX_NORM"
        ],
        "설명": [
            "원재료 및 환율 변동에 따른 가격 부담 수준을 반영한다.",
            "공급집중과 공급국 위험이 결합된 구조적 수급 취약성을 반영한다.",
            "글로벌 공급망 혼잡과 물류 병목 수준을 반영한다.",
            "통상정책, 규제, 지정학적 이벤트에 따른 정책 충격 노출도를 반영한다."
        ]
    })
    st.dataframe(risk_basis_df, use_container_width=True)

    st.markdown("#### 체인별 최종위험 추이")
    st.caption("연월별 최종위험점수의 변화를 통해 특정 체인이 구조적으로 불안정한지, 특정 시기에만 급등했는지를 함께 확인할 수 있다.")
    fig = px.line(
        panel.sort_values("연월"),
        x="연월",
        y="최종위험점수",
        color="체인구분",
        markers=True,
        template="plotly_dark"
    )
    fig.update_layout(height=430, legend_title_text="체인구분")
    st.plotly_chart(fig, use_container_width=True)

    axis_cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in month_df.columns]
    if axis_cols:
        st.markdown("#### 기준 연월 4대 리스크 축 비교")
        st.caption("동일 연월에서 체인별 위험이 어떤 축에서 높게 형성되었는지를 비교한다.")
        melt_df = month_df.melt(id_vars=["체인구분"], value_vars=axis_cols, var_name="리스크축", value_name="점수")
        fig2 = px.bar(
            melt_df,
            x="리스크축",
            y="점수",
            color="체인구분",
            barmode="group",
            template="plotly_dark"
        )
        fig2.update_layout(height=420, legend_title_text="체인구분")
        st.plotly_chart(fig2, use_container_width=True)

    if chain_compare is not None:
        st.markdown("#### 체인별 비교표")
        st.dataframe(chain_compare, use_container_width=True)
        if alert is not None and "우선관리대상_비중" in chain_compare.columns:
            for _, r in chain_compare.iterrows():
                try:
                    v = float(r["우선관리대상_비중"])
                except Exception:
                    continue
                if v == 0:
                    msg = get_priority_zero_reason(alert, r["체인구분"])
                    if msg:
                        st.caption(f"※ {msg}")

# =========================================================
# 2. 체인별 심층 분석
# =========================================================
elif menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")
    st.markdown(
        """
**분석 목적**  
특정 체인의 위험 수준이 시계열상 어떻게 변화해 왔는지와 구조적 취약성이 무엇인지를 분석하기 위한 화면이다.

**핵심 구성**  
1. 월별 최종위험점수 및 4대 리스크 축 추이  
2. 수입구조와 집중도 지표의 장기 흐름 확인  
3. 선택 연월 기준 핵심 구조지표 점검

**해석 기준**  
단기 급등 여부보다, 특정 체인이 장기간 어떤 구조적 위험을 내포하고 있는지 판단하는 데 초점을 둔다.
"""
    )

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel))
    sub = panel[panel["체인구분"] == chain].sort_values("연월").copy()
    months = get_month_list(sub)
    month = st.selectbox("기준 연월 선택", months, index=get_default_index(months))
    row = sub[sub["연월"] == month]
    if row.empty:
        st.warning("선택 연월 데이터가 없습니다.")
        st.stop()
    row = row.iloc[0]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기준 최종위험점수", fmt_num(row.get("최종위험점수", np.nan)), row.get("최종경보등급", "-"))
    c2.metric("상위1국의존도", fmt_pct(row.get("상위1국의존도", np.nan)))
    c3.metric("HHI", fmt_num(row.get("HHI", np.nan)))
    c4.metric("FTA 비중", fmt_pct(row.get("fta_ratio", np.nan)))
    st.caption("상단 지표는 선택한 기준 연월 값을 보여준다.")

    plot_cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수"] if c in sub.columns]
    melt = sub.melt(id_vars=["연월"], value_vars=plot_cols, var_name="지표", value_name="값")
    fig = px.line(melt, x="연월", y="값", color="지표", markers=True, template="plotly_dark")
    fig.update_layout(height=460, legend_title_text="지표")
    st.plotly_chart(fig, use_container_width=True)

    display_cols = [c for c in ["연월", "총수입금액", "총수입물량", "평균수입단가", "수입국수", "지역권수", "상위1국의존도", "상위3국집중도", "HHI", "경보점수기초", "국가보정합계", "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "최종위험점수", "최종경보등급"] if c in sub.columns]
    st.markdown("#### 월별 상세 데이터")
    st.dataframe(sub[display_cols], use_container_width=True)

# =========================================================
# 3. 국가/공급선 상세 분석
# =========================================================
elif menu == "3. 국가/공급선 상세 분석":
    st.header("3. 국가/공급선 상세 분석")
    st.markdown(
        """
**분석 목적**  
특정 체인과 연월에서 어떤 국가가 공급망 위험에 실질적으로 영향을 주고 있는지 식별하기 위한 화면이다.

**핵심 구성**  
1. 국가별 수입비중 비교  
2. 국가별 최종보정점수 비교  
3. 수입비중과 최종보정점수의 결합 분석

**해석 기준**  
수입비중이 높고 최종보정점수도 높은 국가는 우선 점검 대상이며, 비중은 낮지만 점수가 높은 국가는 잠재 위험국으로 해석할 수 있다.
"""
    )

    if country is None:
        st.error("COUNTRY_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(country), key="country_chain")
    months = get_month_list(country[country["체인구분"] == chain])
    month = st.selectbox("연월 선택", months, index=get_default_index(months), key="country_month")
    sub = country[(country["체인구분"] == chain) & (country["연월"] == month)].copy()

    if sub.empty:
        st.warning("선택 조건 데이터가 없습니다.")
        st.stop()

    sort_cols = [c for c in ["금액비중", "국가수입금액"] if c in sub.columns]
    if sort_cols:
        sub = sub.sort_values(sort_cols, ascending=False)

    c1, c2, c3 = st.columns(3)
    c1.metric("공급국 수", sub["국가명"].nunique() if "국가명" in sub.columns else len(sub))
    c2.metric("평균 최종보정점수", fmt_num(safe_numeric(sub["최종보정점수"]).mean()) if "최종보정점수" in sub.columns else "-")
    c3.metric("상위공급국 수", int((sub["상위공급국여부"] == "Y").sum()) if "상위공급국여부" in sub.columns else 0)

    topn = st.slider("상위 몇 개 국가를 볼지", 5, min(30, len(sub)) if len(sub) > 0 else 5, min(10, len(sub)) if len(sub) > 0 else 5)
    top = sub.head(topn).copy()

    st.markdown("#### ① 국가별 금액비중")
    st.caption("해당 체인·연월에서 어느 국가가 실제 수입구조를 주도하는지 보여준다. 비중이 높은 국가는 공급망 영향력이 크므로 우선 점검 필요성이 높다.")
    if {"국가명", "금액비중"}.issubset(top.columns):
        fig = px.bar(
            top.sort_values("금액비중", ascending=False),
            x="국가명",
            y="금액비중",
            color="최종판정" if "최종판정" in top.columns else None,
            color_discrete_map=RISK_COLOR_MAP,
            template="plotly_dark",
            hover_data=[c for c in ["FTA여부", "상위공급국여부", "최종보정점수", "국가수입금액"] if c in top.columns]
        )
        fig.update_layout(height=420, legend_title_text="최종판정")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### ② 국가별 최종보정점수")
    st.caption("국가별 조달 리스크 수준을 직접 비교하기 위한 그래프다. 최종보정점수는 국가기본위험에 공급집중, 지역집중, 상위공급국 여부, FTA 여부 등의 보정요인을 반영한 값이다.")
    if {"국가명", "최종보정점수", "FTA여부"}.issubset(top.columns):
        fig_score = px.bar(
            top.sort_values("최종보정점수", ascending=False),
            x="국가명",
            y="최종보정점수",
            color="FTA여부",
            color_discrete_map=FTA_COLOR_MAP,
            template="plotly_dark",
            hover_data=[c for c in ["금액비중", "상위공급국여부", "최종판정", "국가수입금액"] if c in top.columns]
        )
        fig_score.update_layout(height=420, legend_title_text="FTA 여부")
        st.plotly_chart(fig_score, use_container_width=True)

    st.markdown("#### ③ 금액비중과 최종보정점수의 결합 비교")
    st.caption("수입비중과 리스크 수준을 동시에 비교하기 위한 그래프다. 오른쪽 위 국가는 의존도도 높고 리스크도 높은 국가이므로 우선 점검 대상으로 해석할 수 있다.")
    if {"금액비중", "최종보정점수", "국가명"}.issubset(sub.columns):
        fig2 = px.scatter(
            sub,
            x="금액비중",
            y="최종보정점수",
            size="국가수입금액" if "국가수입금액" in sub.columns else None,
            color="FTA여부" if "FTA여부" in sub.columns else None,
            symbol="상위공급국여부" if "상위공급국여부" in sub.columns else None,
            color_discrete_map=FTA_COLOR_MAP,
            hover_name="국가명",
            hover_data=[c for c in ["상위공급국여부", "최종판정", "금액비중", "최종보정점수"] if c in sub.columns],
            template="plotly_dark"
        )
        fig2.update_layout(height=440, legend_title_text="FTA 여부 / 상위공급국")
        st.plotly_chart(fig2, use_container_width=True)

    show_cols = [c for c in ["연월", "국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부", "국가수입금액", "금액비중", "기본평가점수", "공급국집중보정", "지역권집중보정", "HHI보정", "상위공급국보정", "FTA보정", "총보정점수", "최종보정점수", "최종판정", "비고"] if c in sub.columns]
    st.dataframe(sub[show_cols], use_container_width=True)

# =========================================================
# 4. 충격 원인 추적
# =========================================================
elif menu == "4. 충격 원인 추적":
    st.header("4. 충격 원인 추적")
    st.markdown(
        """
**분석 목적**  
특정 월의 위험 상승이 어떤 요인에서 비롯되었는지를 직접 해석하기 위한 화면이다.

**핵심 구성**  
1. 가격, 수급, 물류, 정책이벤트의 4개 리스크 축 비교  
2. 선택 연월의 핵심 구조지표와 원천지표 확인  
3. GSCPI, TPU_INDEX 등 외생 변수의 동시 점검

**해석 기준**  
체인별 심층 분석이 장기 구조를 파악하는 데 초점을 둔다면, 본 화면은 특정 시점의 경보 원인을 분해하여 설명하는 데 목적이 있다.
"""
    )

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel), key="shock_chain")
    months = get_month_list(panel[panel["체인구분"] == chain])
    month = st.selectbox("연월 선택", months, index=get_default_index(months), key="shock_month")
    row = panel[(panel["체인구분"] == chain) & (panel["연월"] == month)].copy()
    if row.empty:
        st.warning("선택 데이터가 없습니다.")
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
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan)), row.get("최종경보등급", grade_final_risk(row.get("최종위험점수", np.nan))))
    c2.metric("주요 원인축", axis_df.iloc[0]["리스크축"])
    c3.metric("보정사유", row.get("보정사유", "-"))

    fig = px.bar(axis_df, x="리스크축", y="점수", color="리스크축", template="plotly_dark")
    fig.update_layout(height=390, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 선택 월 원인 해석")
    st.markdown(
        f"1. 가격리스크점수: {fmt_num(row.get('가격리스크점수', np.nan), 2)}  \n"
        f"2. 수급리스크점수: {fmt_num(row.get('수급리스크점수', np.nan), 2)} "
        f"(HHI {fmt_num(row.get('HHI', np.nan), 2)}, 상위1국의존도 {fmt_pct(row.get('상위1국의존도', np.nan), 2)}, 국가보정합계 {fmt_num(row.get('국가보정합계', np.nan), 2)})  \n"
        f"3. 물류리스크점수: {fmt_num(row.get('물류리스크점수', np.nan), 2)} (GSCPI_Norm 기준)  \n"
        f"4. 정책이벤트리스크점수: {fmt_num(row.get('정책이벤트리스크점수', np.nan), 2)} (TPU_INDEX_NORM 기준)"
    )

    if gscpi is not None and "연월" in gscpi.columns:
        gsub = gscpi[gscpi["연월"] == month]
        if not gsub.empty:
            g = gsub.iloc[0]
            st.info(f"GSCPI 원천값: {fmt_num(g.get('GSCPI', np.nan), 2)} / 정규화값: {fmt_num(g.get('GSCPI_Norm', g.get('GSCPI_NORM', np.nan)), 2)}")

    if tpu is not None and "연월" in tpu.columns:
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
elif menu == "5. 선행 신호 후보 탐지":
    st.header("5. 선행 신호 후보 탐지")
    st.markdown(
        """
**분석 목적**  
최종위험점수에 선행하여 움직일 가능성이 있는 원천 지표를 식별하기 위한 화면이다.

**핵심 구성**  
본 화면은 ① 후보 변수 범위, ② 변수별 최적 lag 요약, ③ lag 1~6 상세 비교, ④ 분석 베이스 데이터 확인의 네 단계로 구성된다.

**해석 기준**  
절대상관이 크고 일관된 방향성을 보이는 지표일수록 선행 모니터링 가치가 높다고 해석한다.
"""
    )

    if signal_scope is None or lead_compare is None or lead_detail is None:
        st.error("SIGNAL / LEAD_SIGNAL 관련 시트가 없습니다.")
        st.stop()

    chain_options = get_chain_list(lead_compare)
    chain = st.selectbox("체인 선택", chain_options, key="lead_chain")

    st.markdown("#### ① 후보 변수 범위")
    st.caption("선행신호 검토 대상에 포함된 변수와 제외된 변수를 구분하여 제시한다. 분석대상 선정 범위를 먼저 확인하는 단계이다.")
    scope_sub = signal_scope[signal_scope["체인구분"].isin(["공통", chain])] if "체인구분" in signal_scope.columns else signal_scope.copy()
    st.dataframe(scope_sub, use_container_width=True)

    st.markdown("#### ② 변수별 최적 lag 요약")
    st.caption("각 지표가 lag 1~6개월 중 어느 시점에서 최종위험과 가장 강한 관계를 보였는지 요약한다. 선행 모니터링 우선순위를 정하는 핵심 구간이다.")
    compare_sub = lead_compare[lead_compare["체인구분"] == chain].copy() if "체인구분" in lead_compare.columns else lead_compare.copy()
    if "최적lag절대상관" in compare_sub.columns:
        compare_sub = compare_sub.sort_values("최적lag절대상관", ascending=False)
    st.dataframe(compare_sub, use_container_width=True)

    if {"지표", "최적lag상관"}.issubset(compare_sub.columns):
        fig = px.bar(
            compare_sub,
            x="지표",
            y="최적lag상관",
            color="최종해석" if "최종해석" in compare_sub.columns else None,
            template="plotly_dark",
            hover_data=[c for c in ["최적lag개월", "최적lag절대상관", "최적lag관측치N"] if c in compare_sub.columns]
        )
        fig.update_layout(height=430)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### ③ 지표별 lag 1~6 상세")
    st.caption("특정 지표를 선택하여 월차별 상관 패턴을 세부적으로 확인한다. 단순히 상관이 큰지보다 어느 시점에서 가장 안정적으로 관계가 나타나는지를 함께 본다.")
    metric_list = compare_sub["지표"].dropna().astype(str).tolist() if "지표" in compare_sub.columns else []
    if metric_list:
        metric = st.selectbox("지표 선택", metric_list, key="lead_metric")
        detail_sub = lead_detail[(lead_detail["체인구분"] == chain) & (lead_detail["지표"] == metric)].copy() if {"체인구분", "지표"}.issubset(lead_detail.columns) else lead_detail.copy()
        if "lag개월" in detail_sub.columns:
            detail_sub = detail_sub.sort_values("lag개월")
        st.dataframe(detail_sub, use_container_width=True)
        if {"lag개월", "상관계수"}.issubset(detail_sub.columns):
            fig2 = px.line(detail_sub, x="lag개월", y="상관계수", markers=True, template="plotly_dark")
            fig2.update_layout(height=380, xaxis=dict(dtick=1))
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("#### ④ 분석 베이스 데이터 확인")
        st.caption("실제 분석에 사용된 시계열을 확인함으로써 결과 해석의 근거를 검토할 수 있다.")
        if signal_base is not None:
            cols = [c for c in ["연월", "체인구분", "최종위험점수_V2", metric] if c in signal_base.columns]
            base_sub = signal_base[signal_base["체인구분"] == chain][cols].copy() if "체인구분" in signal_base.columns else signal_base[cols].copy()
            st.dataframe(base_sub, use_container_width=True)

# =========================================================
# 6. 기업 대응 우선순위 추천 / 시뮬레이터
# =========================================================
elif menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.header("6. 기업 대응 우선순위 추천 / 시뮬레이터")
    st.markdown(
        """
**분석 목적**  
현재 공급구조에서 어떤 대응 과제를 우선적으로 검토해야 하는지를 구조적으로 제시하기 위한 화면이다.

**핵심 구성**  
1. 구조 진단과 우선 개선 과제 제시  
2. 시나리오형 민감도 분석  
3. 직접 조정형 시뮬레이션을 통한 대응방안 비교

**해석 기준**  
최종위험점수 자체보다, 어떤 구조요인을 조정할 때 위험 완화 효과가 상대적으로 크게 나타나는지를 비교하는 데 의미가 있다.
"""
    )

    if panel is None:
        st.error("PANEL_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(panel), key="prio_chain")
    months = get_month_list(panel[panel["체인구분"] == chain])
    month = st.selectbox("연월 선택", months, index=get_default_index(months), key="prio_month")
    row_df = panel[(panel["체인구분"] == chain) & (panel["연월"] == month)].copy()
    if row_df.empty:
        st.warning("선택 데이터가 없습니다.")
        st.stop()
    row = row_df.iloc[0].to_dict()

    if alert is not None and {"연월", "체인구분"}.issubset(alert.columns):
        a = alert[(alert["체인구분"] == chain) & (alert["연월"] == month)]
        if not a.empty:
            row.update(a.iloc[0].to_dict())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan), 2), row.get("최종경보등급", "-"))
    c2.metric("대체조달가능성", fmt_num(row.get("대체조달가능성_점수", np.nan), 2), row.get("대체조달가능성", "-"))
    c3.metric("보정사유", row.get("보정사유", "-"))
    c4.metric("우선관리대상", row.get("상대적_우선관리대상", "-"))

    st.markdown("#### 경영진 요약 코멘트")
    st.write(build_executive_comment(row))

    st.markdown("#### 구조 진단")
    diag = pd.DataFrame({
        "항목": ["상위1국의존도", "HHI", "FTA 비중", "지역권수", "수입국수"],
        "현재값": [row.get("상위1국의존도", np.nan), row.get("HHI", np.nan), row.get("fta_ratio", np.nan), row.get("지역권수", np.nan), row.get("수입국수", np.nan)]
    })
    st.dataframe(diag, use_container_width=True)

    st.markdown("#### 권고 액션 로드맵")
    st.dataframe(build_action_roadmap(row), use_container_width=True)

    st.markdown("#### 시나리오형 시뮬레이터")
    scenario = st.selectbox("대표 시나리오 선택", ["현 상태 유지", "보수적 분산", "공격적 분산", "FTA 확대", "물류 완화", "정책 충격 완화"])

    st.caption("시나리오 정의")
    for idx, line in enumerate(scenario_rule_text(scenario), start=1):
        st.markdown(f"{idx}. {line}")
    st.caption("주의: 본 시뮬레이션은 예측계수 기반 전망이 아니라, 관측 가능한 구조지표를 기준으로 변화폭을 규칙화한 민감도 분석 결과이다. 실제 효과는 계약 구조, 시장 상황, 조달 실행력에 따라 달라질 수 있다.")

    current_est = estimate_final_score(panel, row, entropy)
    scenario_row = apply_scenario_to_row(row, scenario) if scenario != "현 상태 유지" else dict(row)
    scenario_est = estimate_final_score(panel, scenario_row, entropy)

    sim_df = pd.DataFrame({
        "항목": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수", "상위1국의존도", "HHI", "FTA 비중", "수입국수", "지역권수"],
        "현재": [row.get("가격리스크점수", np.nan), row.get("수급리스크점수", np.nan), row.get("물류리스크점수", np.nan), row.get("정책이벤트리스크점수", np.nan), row.get("상위1국의존도", np.nan), row.get("HHI", np.nan), row.get("fta_ratio", np.nan), row.get("수입국수", np.nan), row.get("지역권수", np.nan)],
        "시나리오 적용 후": [scenario_row.get("가격리스크점수", np.nan), scenario_row.get("수급리스크점수", np.nan), scenario_row.get("물류리스크점수", np.nan), scenario_row.get("정책이벤트리스크점수", np.nan), scenario_row.get("상위1국의존도", np.nan), scenario_row.get("HHI", np.nan), scenario_row.get("fta_ratio", np.nan), scenario_row.get("수입국수", np.nan), scenario_row.get("지역권수", np.nan)],
    })
    st.dataframe(sim_df, use_container_width=True)

    s1, s2 = st.columns(2)
    s1.metric("현재 추정 최종위험점수", fmt_num(current_est, 2))
    s2.metric("시나리오 적용 후 추정 최종위험점수", fmt_num(scenario_est, 2), delta=fmt_num(scenario_est - current_est, 2) if pd.notna(current_est) and pd.notna(scenario_est) else "-")

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
elif menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")
    st.markdown(
        """
**분석 목적**  
기존 공급구조를 보완하거나 대체할 수 있는 공급국 후보를 식별하기 위한 화면이다.

**핵심 구성**  
1. 국가별 추천점수 정렬  
2. 최종보정점수와 FTA 여부의 동시 비교  
3. 신규 보완 공급선 검토를 위한 필터링 제공

**해석 기준**  
추천점수가 높다고 해서 즉시 최적 대체국을 의미하는 것은 아니며, 위험 수준, FTA 이점, 기존 조달 실적을 함께 고려한 후보군으로 해석한다.
"""
    )

    if country is None:
        st.error("COUNTRY_MONTHLY 시트가 없습니다.")
        st.stop()

    chain = st.selectbox("체인 선택", get_chain_list(country), key="alt_chain")
    months = get_month_list(country[country["체인구분"] == chain])
    month = st.selectbox("연월 선택", months, index=get_default_index(months), key="alt_month")

    cur = country[(country["체인구분"] == chain) & (country["연월"] == month)].copy()
    if cur.empty:
        st.warning("선택 데이터가 없습니다.")
        st.stop()

    fta_only = st.checkbox("FTA 체결국만 보기", value=False)
    exclude_top = st.checkbox("상위공급국 제외", value=False, help="기존 주력 공급국이 아니라 신규 보완 후보를 찾고 싶을 때 사용한다.")

    if fta_only and "FTA여부" in cur.columns:
        cur = cur[cur["FTA여부"] == "Y"].copy()
    if exclude_top and "상위공급국여부" in cur.columns:
        cur = cur[cur["상위공급국여부"] != "Y"].copy()

    cur["FTA가점"] = np.where(cur["FTA여부"] == "Y", 15, 0) if "FTA여부" in cur.columns else 0
    cur["비상위가점"] = np.where(cur["상위공급국여부"] == "N", 10, 0) if "상위공급국여부" in cur.columns else 0

    if "최종보정점수" in cur.columns:
        cur["안정성점수"] = 100 - safe_numeric(cur["최종보정점수"])
    elif "기본평가점수" in cur.columns:
        cur["안정성점수"] = 100 - safe_numeric(cur["기본평가점수"])
    else:
        cur["안정성점수"] = 0

    if "국가수입금액" in cur.columns:
        cur["규모가점"] = safe_numeric(cur["국가수입금액"]).rank(pct=True) * 10
    else:
        cur["규모가점"] = 0

    cur["대체국추천점수"] = cur["안정성점수"].fillna(0) + cur["FTA가점"] + cur["비상위가점"] + cur["규모가점"]
    cur = cur.sort_values("대체국추천점수", ascending=False)

    show_cols = [c for c in ["국가코드", "국가명", "지역권", "FTA여부", "상위공급국여부", "국가수입금액", "금액비중", "기본평가점수", "최종보정점수", "최종판정", "대체국추천점수"] if c in cur.columns]
    st.dataframe(cur[show_cols], use_container_width=True)

    top10 = cur.head(10).copy()
    if {"국가명", "대체국추천점수"}.issubset(top10.columns):
        fig = px.bar(
            top10.sort_values("대체국추천점수", ascending=False),
            x="국가명",
            y="대체국추천점수",
            color="최종판정" if "최종판정" in top10.columns else None,
            color_discrete_map=RISK_COLOR_MAP,
            template="plotly_dark",
            hover_data=[c for c in ["FTA여부", "상위공급국여부", "최종보정점수"] if c in top10.columns]
        )
        fig.update_layout(height=420, legend_title_text="최종판정")
        st.plotly_chart(fig, use_container_width=True)
    st.caption("색상은 추천 후보의 위험 판정 수준을 함께 보여주기 위해 최종판정 기준으로 구분하였다.")

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")
    st.markdown(
        """
**분석 목적**  
대시보드에 연결된 실제 원천 시트와 데이터를 직접 점검하기 위한 화면이다.

**핵심 구성**  
1. 시트별 원본 테이블 확인  
2. 행·열 규모 확인  
3. CSV 및 Excel 다운로드 제공

**해석 기준**  
결과 검증, 후속 분석, 제출 전 교차점검을 위한 기초 데이터 확인 단계로 사용한다.
"""
    )

    available = {k: v for k, v in sheet_map.items() if v is not None}
    labels = [f"{v} ({k})" for k, v in available.items()]
    label = st.selectbox("시트 선택", labels)
    key = {f"{v} ({k})": k for k, v in available.items()}[label]
    df = data[key]

    st.write(f"선택 시트: **{available[key]}**")
    st.write(f"행 수: {len(df):,} / 열 수: {len(df.columns):,}")
    st.dataframe(df.head(200), use_container_width=True)
    show_df_download_button(df, "현재 시트 CSV 다운로드", f"{available[key]}.csv")

    excel_bytes = make_download_excel({sheet_map[k]: data[k] for k in available.keys() if data.get(k) is not None})
    st.download_button(
        "로드된 전체 시트 Excel로 다시 다운로드",
        data=excel_bytes,
        file_name="dashboard_loaded_workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================================================
# 9. 데이터 검증 / 방법론
# =========================================================
elif menu == "9. 데이터 검증 / 방법론":
    st.header("9. 데이터 검증 / 방법론")
    st.markdown(
        """
**분석 목적**  
데이터 연결과 계산 구조가 정상적으로 작동하는지 점검하기 위한 화면이다.

**핵심 구성**  
1. 필수 시트 존재 여부 확인  
2. 핵심 컬럼 존재 여부 점검  
3. PANEL/ALERT 정합성 및 lag 구조 검증

**해석 기준**  
분석 결과 자체보다, 결과를 뒷받침하는 데이터 구조와 계산 일관성을 확인하는 검증 단계로 해석한다.
"""
    )

    st.markdown("#### 데이터 검증 결과")
    st.dataframe(check_df, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("PASS", int((check_df["레벨"] == "PASS").sum()))
    c2.metric("WARN", int((check_df["레벨"] == "WARN").sum()))
    c3.metric("FAIL", int((check_df["레벨"] == "FAIL").sum()))

    st.markdown("#### 시트 매핑")
    st.dataframe(pd.DataFrame({"logical_key": list(sheet_map.keys()), "loaded_sheet": list(sheet_map.values())}), use_container_width=True)

    if entropy is not None:
        st.markdown("#### ENTROPY_WEIGHT")
        st.dataframe(entropy, use_container_width=True)
    if norm_check is not None:
        st.markdown("#### NOMALIZATION_CHECK")
        st.dataframe(norm_check, use_container_width=True)
    if norm_audit is not None:
        st.markdown("#### NOMALIZATION_AUDIT")
        st.dataframe(norm_audit, use_container_width=True)
    if method is not None:
        st.markdown("#### METHOD_GUIDE")
        st.dataframe(method, use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.caption("팀명: 망보는사람들")
st.sidebar.caption("최종 시트 기반 공급망 리스크 분석 앱")
