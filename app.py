import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(
    page_title="망보는사람들 공급망 리스크 대시보드",
    page_icon="📊",
    layout="wide"
)

CHAIN_ORDER = ["납산배터리군", "리튬이온배터리군"]
CHAIN_COLOR_MAP = {
    "납산배터리군": "#1f77b4",
    "리튬이온배터리군": "#d62728",
}
GRADE_COLOR_MAP = {
    "낮음": "#4C78A8",
    "보통": "#F2C14E",
    "높음": "#F28E2B",
    "매우높음": "#E15759",
    "해석유보": "#9D9D9D",
}
AXIS_COLOR_MAP = {
    "가격리스크점수": "#4C78A8",
    "수급리스크점수": "#E15759",
    "물류리스크점수": "#59A14F",
    "정책이벤트리스크점수": "#B07AA1",
}
LAG_COLOR_MAP = {
    "1": "#08306B",
    "2": "#2171B5",
    "3": "#4292C6",
    "4": "#6BAED6",
    "5": "#9ECAE1",
    "6": "#C6DBEF",
}

SHEET_DESC = {
    "COUNTRY_MONTHLY": "체인·연월별 국가 단위 수입 비중, 기본평가점수, 최종보정점수, FTA 여부, 공급선 판정 정보를 담은 시트입니다. 국가/공급선 상세분석과 대체국 추천 메뉴의 핵심 원천입니다.",
    "PANEL_MONTHLY": "체인·연월별 핵심 집계 시트입니다. 가격리스크점수, 수급리스크점수, 물류리스크점수, 정책이벤트리스크점수, 최종위험점수(원점수), 상대위험지수(정규화값)가 들어 있습니다.",
    "ALERT_RESULT": "체인·연월별 최종 경보 결과 시트입니다. 최종경보등급, 대체조달가능성 점수, 상대적 우선관리대상, 보정사유, 비고 등을 담고 있습니다.",
    "체인별 비교표": "체인별 평균 수준, 분위수(Q25·Q50·Q75), 우선관리 비중, 30점 이상 개월수 등 비교용 요약값을 정리한 시트입니다.",
    "ENTROPY_WEIGHT": "체인별 엔트로피 가중치(EWM) 시트입니다. 가격, 수급기초, 수급최종, 최종결합, 대체조달가능성 가중치가 들어 있습니다.",
    "NOMALIZATION_CHECK": "정규화 및 분위수 계산 결과를 점검하기 위한 검증 시트입니다. min, q25, median, q75, max가 정리되어 있습니다.",
    "NOMALIZATION_AUDIT": "정규화 결과를 더 세부적으로 확인하기 위한 감사용 시트입니다.",
    "METHOD_GUIDE": "데이터 처리 절차, 체인별 계산 구조, 가중치 결합 방식 등 방법론 설명 시트입니다.",
    "SIGNAL_BASE": "선행신호 분석용 원천 변수 시트입니다. 환율, 원재료 가격, HHI, 상위1국의존도, 수입국수, 지역권수 등 후보변수가 들어 있습니다.",
    "SIGNAL_LAG_TABLE": "선행신호 후보변수의 lag 1~6개월 시차 테이블입니다.",
    "LEAD_SIGNAL_LAG_DETAIL": "체인·지표·lag별 상관계수와 관측치 수를 담은 상세 시트입니다.",
    "LEAD_SIGNAL_LAG_COMPARE": "지표별 최적 lag와 절대상관을 요약 비교한 시트입니다.",
    "HS_MONTHLY_SUMMARY": "HS코드 기준 월별 집계 시트입니다. 품목 수준 구조를 보조적으로 확인할 때 사용합니다.",
    "MARKET_INDEX": "환율, 납가격, 리튬가격, 니켈가격 등 시장지표의 월별 원천 시트입니다. 가격리스크점수의 기반이 됩니다.",
    "MAREKT_INDEX": "환율, 납가격, 리튬가격, 니켈가격 등 시장지표의 월별 원천 시트입니다. 가격리스크점수의 기반이 됩니다.",
    "GSCPI_INDEX": "글로벌 공급망 압력지수(GSCPI)의 월별 시트입니다. 물류리스크점수의 원천입니다.",
    "TPU_INDEX": "무역정책 불확실성 관련 지표(TPU_INDEX)의 월별 시트입니다. 정책이벤트리스크점수의 원천입니다.",
}

# ---------------------------------------------------------
# 공통 함수
# ---------------------------------------------------------
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce")

def fmt_num(v, digits=2):
    if pd.isna(v):
        return "-"
    return f"{float(v):,.{digits}f}"

def fmt_pct(v, digits=2):
    if pd.isna(v):
        return "-"
    return f"{float(v):,.{digits}f}%"

def safe_ym(x, year=None, month=None):
    if pd.notna(year) and pd.notna(month):
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
                    continue
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
        return s

    return None

def ensure_ym_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols = df.columns.tolist()

    if "연월_키" in cols:
        dt = pd.to_datetime(df["연월_키"], errors="coerce")
        if dt.notna().sum() > 0:
            df["연월"] = dt.dt.strftime("%Y-%m")
            return df

    if "연도" in cols and "월" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r.get("연도"), r.get("월")), axis=1)
        return df

    if "Year" in cols and "Month" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r.get("Year"), r.get("Month")), axis=1)
        return df

    if "연" in cols and "월" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r.get("연"), r.get("월")), axis=1)
        return df

    if "연월" in cols:
        df["연월"] = df["연월"].apply(safe_ym)
        return df

    return df

def get_month_list(df):
    if df is None or df.empty or "연월" not in df.columns:
        return []
    vals = df["연월"].dropna().astype(str).unique().tolist()
    vals = [v for v in vals if len(v) == 7 and v[4] == "-"]
    return sorted(vals)

def get_chain_list(df):
    if df is None or df.empty or "체인구분" not in df.columns:
        return []
    vals = df["체인구분"].dropna().astype(str).str.strip().unique().tolist()
    return [x for x in CHAIN_ORDER if x in vals] + [x for x in vals if x not in CHAIN_ORDER]

def info_box(title, bullets):
    if isinstance(bullets, str):
        bullets = [bullets]
    li = "".join([f"<li>{b}</li>" for b in bullets])
    st.markdown(
        f"""
        <div style="border:1px solid rgba(128,128,128,0.25); border-radius:12px; padding:14px 16px; margin-bottom:14px;">
            <div style="font-weight:700; margin-bottom:8px;">{title}</div>
            <ul style="margin-top:0; padding-left:18px; line-height:1.7; font-size:0.96rem;">{li}</ul>
        </div>
        """,
        unsafe_allow_html=True
    )

def style_metric_container():
    st.markdown("""
    <style>
    div[data-testid="stMetric"] {
        background-color: rgba(255,255,255,0.04);
        border: 1px solid rgba(128,128,128,0.25);
        padding: 12px 14px;
        border-radius: 12px;
    }
    </style>
    """, unsafe_allow_html=True)

def month_chain_slice(df, month=None, chain=None):
    tmp = df.copy()
    if month and "연월" in tmp.columns:
        tmp = tmp[tmp["연월"] == month].copy()
    if chain and "체인구분" in tmp.columns:
        tmp = tmp[tmp["체인구분"] == chain].copy()
    return tmp

def risk_label_from_0_100(score):
    if pd.isna(score):
        return "해석유보"
    if score >= 75:
        return "매우높음"
    if score >= 50:
        return "높음"
    if score >= 25:
        return "보통"
    return "낮음"

def build_download_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=str(sname)[:31], index=False)
    return output.getvalue()

def aggregate_country_view(country_df: pd.DataFrame) -> pd.DataFrame:
    if country_df is None or country_df.empty:
        return pd.DataFrame()

    tmp = country_df.copy()

    if "국가별수입비중" not in tmp.columns:
        for fallback in ["금액비중", "중량비중"]:
            if fallback in tmp.columns:
                tmp["국가별수입비중"] = safe_numeric(tmp[fallback])
                break

    if "국가수입금액" in tmp.columns:
        tmp["_w"] = safe_numeric(tmp["국가수입금액"]).fillna(0)
    else:
        tmp["_w"] = safe_numeric(tmp.get("국가별수입비중", 0)).fillna(0)

    group_cols = [c for c in ["국가코드", "국가명", "지역권"] if c in tmp.columns]
    if not group_cols:
        group_cols = ["국가명"] if "국가명" in tmp.columns else []

    rows = []
    for keys, g in tmp.groupby(group_cols, dropna=False):
        if not isinstance(keys, tuple):
            keys = (keys,)
        row = {}
        for c, v in zip(group_cols, keys):
            row[c] = v

        for c in ["국가별수입비중", "국가수입금액", "국가수입중량", "금액비중", "중량비중", "지역권별수입비중"]:
            if c in g.columns:
                row[c] = safe_numeric(g[c]).sum()

        for c in ["기본평가점수", "최종보정점수"]:
            if c in g.columns:
                weights = np.maximum(g["_w"].fillna(0), 1e-9)
                row[c] = np.average(safe_numeric(g[c]).fillna(0), weights=weights)

        row["FTA여부"] = "Y" if ("FTA여부" in g.columns and (g["FTA여부"].astype(str) == "Y").any()) else "N"
        row["상위공급국여부"] = "Y" if ("상위공급국여부" in g.columns and (g["상위공급국여부"].astype(str) == "Y").any()) else "N"

        if "비고" in g.columns:
            vals = g["비고"].dropna().astype(str).unique().tolist()
            row["비고"] = ", ".join(vals[:2])

        row["국가기본위험_평가등급"] = risk_label_from_0_100(row.get("기본평가점수", np.nan))
        row["국가공급선_최종판정"] = risk_label_from_0_100(row.get("최종보정점수", np.nan))
        rows.append(row)

    out = pd.DataFrame(rows)
    if "국가별수입비중" in out.columns:
        out = out.sort_values("국가별수입비중", ascending=False).reset_index(drop=True)
    return out

def get_processed_sheet_for_display(selected_sheet, loaded):
    name_map = {
        "COUNTRY_MONTHLY": loaded.get("country", pd.DataFrame()),
        "PANEL_MONTHLY": loaded.get("panel", pd.DataFrame()),
        "ALERT_RESULT": loaded.get("alert", pd.DataFrame()),
        "체인별 비교표": loaded.get("compare", pd.DataFrame()),
        "ENTROPY_WEIGHT": loaded.get("entropy", pd.DataFrame()),
        "NOMALIZATION_CHECK": loaded.get("norm_check", pd.DataFrame()),
        "NOMALIZATION_AUDIT": loaded.get("norm_audit", pd.DataFrame()),
        "METHOD_GUIDE": loaded.get("method", pd.DataFrame()),
        "SIGNAL_BASE": loaded.get("signal_base", pd.DataFrame()),
        "SIGNAL_LAG_TABLE": loaded.get("signal_lag", pd.DataFrame()),
        "LEAD_SIGNAL_LAG_DETAIL": loaded.get("lead_detail", pd.DataFrame()),
        "LEAD_SIGNAL_LAG_COMPARE": loaded.get("lead_compare", pd.DataFrame()),
        "HS_MONTHLY_SUMMARY": loaded.get("hs_summary", pd.DataFrame()),
        "MARKET_INDEX": loaded.get("market", pd.DataFrame()),
        "MAREKT_INDEX": loaded.get("market", pd.DataFrame()),
        "GSCPI_INDEX": loaded.get("gscpi", pd.DataFrame()),
        "TPU_INDEX": loaded.get("tpu", pd.DataFrame()),
    }
    return name_map.get(selected_sheet, pd.DataFrame())

def get_weight_map(entropy_df, stage, chain):
    if entropy_df is None or entropy_df.empty:
        return {}
    tmp = entropy_df[(entropy_df["단계"] == stage) & (entropy_df["체인구분"] == chain)].copy()
    if tmp.empty:
        return {}
    return dict(zip(tmp["변수명"], safe_numeric(tmp["가중치"])))

def add_priority_metrics(alert_df, panel_df, compare_df):
    df = alert_df.copy()

    if "최종위험점수(원점수)" not in df.columns and "최종위험점수(원점수)" in panel_df.columns:
        df = df.merge(
            panel_df[["연월", "체인구분", "최종위험점수(원점수)"]],
            on=["연월", "체인구분"], how="left"
        )

    keep = [
        "체인구분",
        "상대위험지수_Q25", "상대위험지수_Q50", "상대위험지수_Q75",
        "최종위험점수(원점수)_Q25", "최종위험점수(원점수)_Q50", "최종위험점수(원점수)_Q75",
    ]
    keep = [c for c in keep if c in compare_df.columns]
    if keep:
        df = df.merge(compare_df[keep], on="체인구분", how="left")

    alt_q = (
        df.groupby("체인구분")["대체조달가능성_점수"]
        .quantile(0.25)
        .rename("대체조달가능성_Q25")
        .reset_index()
    )
    df = df.merge(alt_q, on="체인구분", how="left")

    df["Q75_초과폭"] = safe_numeric(df["상대위험지수(정규화값)"]) - safe_numeric(df["상대위험지수_Q75"])
    df["Q25_미달폭"] = safe_numeric(df["대체조달가능성_Q25"]) - safe_numeric(df["대체조달가능성_점수"])

    # 관리강도는 우선관리대상(Y)인 경우에만 의미를 갖도록 조정
    df["우선관리강도"] = np.where(
        df["상대적_우선관리대상"].astype(str) == "Y",
        df[["Q75_초과폭", "Q25_미달폭"]].clip(lower=0).sum(axis=1),
        0
    )

    return df

def metric_help_text(metric_name, row):
    priority = str(row.get("상대적_우선관리대상", ""))
    q75_gap = row.get("Q75_초과폭", np.nan)
    q25_gap = row.get("Q25_미달폭", np.nan)
    alt_score = row.get("대체조달가능성_점수", np.nan)
    alt_q25 = row.get("대체조달가능성_Q25", np.nan)

    if metric_name == "우선관리대상":
        if priority == "Y":
            return "Y = 고위험·저대체조달 동시 충족"
        return "N = 두 조건을 동시에 충족하지 않음"

    if metric_name == "관리강도":
        if priority == "Y":
            return "산식: max(Q75초과폭,0)+max(Q25미달폭,0)"
        return "N이면 관리강도 0으로 처리"

    if metric_name == "Q75초과폭":
        if pd.notna(q75_gap) and q75_gap > 0:
            return "현재 상대위험지수가 체인 상위구간을 초과"
        return "체인 상위구간(Q75) 이내"

    if metric_name == "Q25미달폭":
        if pd.notna(q25_gap) and q25_gap > 0:
            return f"대체조달가능성 {fmt_num(alt_score)}가 Q25({fmt_num(alt_q25)})보다 낮음"
        return f"대체조달가능성 {fmt_num(alt_score)}가 Q25({fmt_num(alt_q25)}) 이상"

    return ""

def build_quadrant_annotations(summary):
    anns = []
    if summary.empty:
        return anns

    x_min = summary["대체조달가능성_점수"].min()
    x_max = summary["대체조달가능성_점수"].max()
    y_min = summary["상대위험지수(정규화값)"].min()
    y_max = summary["상대위험지수(정규화값)"].max()

    x_mid = (x_min + x_max) / 2
    y_mid = (y_min + y_max) / 2

    anns.append(dict(x=x_min, y=y_max + 4, text="좌상단 = Y 후보 영역(고위험·저대체조달)", showarrow=False, xanchor="left"))
    anns.append(dict(x=x_max, y=y_max + 4, text="우상단 = 위험 높으나 대체조달은 상대적으로 양호", showarrow=False, xanchor="right"))
    anns.append(dict(x=x_min, y=y_min - 4, text="좌하단 = 위험은 낮으나 대체조달 취약", showarrow=False, xanchor="left"))
    anns.append(dict(x=x_max, y=y_min - 4, text="우하단 = 상대적 안정 영역", showarrow=False, xanchor="right"))
    return anns

def get_driver_axis(panel_row):
    axis_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    vals = {c: safe_numeric(pd.Series([panel_row.get(c, np.nan)])).iloc[0] for c in axis_cols}
    vals = {k: v for k, v in vals.items() if pd.notna(v)}
    if not vals:
        return None
    return max(vals, key=vals.get)

def get_tpu_story_text():
    return (
        "TPU_INDEX는 무역정책 불확실성 관련 월별 지표로, 관세·수출통제·제재·통상갈등 등 정책환경 변화가 "
        "공급망에 주는 긴장도를 간접적으로 반영한다. 본 앱에서는 이를 정책이벤트리스크의 원천지표로 사용하며, "
        "값이 높아질수록 정책·지정학 이벤트가 공급망에 부담을 줄 가능성이 커진 것으로 해석한다."
    )

def get_signal_axis(signal_name: str):
    s = str(signal_name)
    if any(k in s for k in ["환율", "납가격", "리튬가격", "니켈가격"]):
        return "가격"
    if any(k in s for k in ["HHI", "상위1국의존도", "수입국수", "지역권수", "fta_ratio", "국가보정합계"]):
        return "수급"
    if any(k in s for k in ["GSCPI"]):
        return "물류"
    if any(k in s for k in ["TPU"]):
        return "정책이벤트"
    return "기타"

def get_signals_for_axis(lead_compare_df, chain, axis_name):
    if lead_compare_df is None or lead_compare_df.empty:
        return pd.DataFrame()
    tmp = lead_compare_df[lead_compare_df["체인구분"] == chain].copy()
    if tmp.empty:
        return tmp
    tmp["축"] = tmp["지표"].astype(str).apply(get_signal_axis)
    tmp = tmp[tmp["축"] == axis_name].copy()
    sort_col = "최적lag절대상관" if "최적lag절대상관" in tmp.columns else tmp.columns[-1]
    return tmp.sort_values(sort_col, ascending=False)

def simulate_intervention(base_row, panel_row, entropy_df, panel_df, lead_compare_df, action_type):
    chain = base_row["체인구분"]
    final_weights = get_weight_map(entropy_df, "PANEL_FINAL", chain)

    axis_map = {
        "가격": "가격리스크점수",
        "수급": "수급리스크점수",
        "물류": "물류리스크점수",
        "정책이벤트": "정책이벤트리스크점수",
    }

    scenarios = {
        "가격 계약구조 조정(장기계약·헤지·판가연동)": {"axis": "가격", "impact_rate": 0.08, "alt_plus": 0, "desc": "가격 충격의 직접 흡수력을 일부 높이는 보수적 가정"},
        "공급선 다변화(보조 공급국 확보·의존도 완화)": {"axis": "수급", "impact_rate": 0.10, "alt_plus": 8, "desc": "상위국 집중도 완화 및 대체가능성 개선을 동시에 반영한 가정"},
        "안전재고·물류경로 다변화": {"axis": "물류", "impact_rate": 0.07, "alt_plus": 3, "desc": "운송 병목 영향 완화를 보수적으로 반영한 가정"},
        "규제모니터링·통관대응·대체시장 사전준비": {"axis": "정책이벤트", "impact_rate": 0.06, "alt_plus": 4, "desc": "정책 충격에 대한 사전 대응체계를 보수적으로 반영한 가정"},
        "FTA 활용국 중심 대체조달 전환": {"axis": "수급", "impact_rate": 0.05, "alt_plus": 10, "desc": "대체조달가능성 개선 효과를 중심으로 반영한 가정"},
    }

    sc = scenarios[action_type]
    target_axis = sc["axis"]
    target_col = axis_map[target_axis]

    old_raw = float(base_row["최종위험점수(원점수)"])
    old_rel = float(base_row["상대위험지수(정규화값)"])
    old_alt = float(base_row["대체조달가능성_점수"])

    axis_val = float(panel_row.get(target_col, 0))
    axis_weight = float(final_weights.get(target_col, 0.25))

    raw_reduction = axis_val * axis_weight * sc["impact_rate"]
    raw_reduction = min(raw_reduction, old_raw * 0.15)  # 과도한 개선 방지
    new_raw = max(0.0, old_raw - raw_reduction)

    chain_panel = panel_df[panel_df["체인구분"] == chain].copy()
    raw_min = float(chain_panel["최종위험점수(원점수)"].min())
    raw_max = float(chain_panel["최종위험점수(원점수)"].max())
    if raw_max != raw_min:
        new_rel = (new_raw - raw_min) / (raw_max - raw_min) * 100
    else:
        new_rel = old_rel
    new_rel = max(0.0, min(100.0, new_rel))

    new_alt = max(0.0, min(100.0, old_alt + sc["alt_plus"]))

    q75 = float(base_row["상대위험지수_Q75"]) if pd.notna(base_row["상대위험지수_Q75"]) else np.nan
    q25 = float(base_row["대체조달가능성_Q25"]) if pd.notna(base_row["대체조달가능성_Q25"]) else np.nan

    new_q75_gap = new_rel - q75 if pd.notna(q75) else np.nan
    new_q25_gap = q25 - new_alt if pd.notna(q25) else np.nan
    new_priority = "Y" if (pd.notna(new_q75_gap) and pd.notna(new_q25_gap) and new_q75_gap >= 0 and new_q25_gap >= 0) else "N"
    new_strength = (max(new_q75_gap, 0) + max(new_q25_gap, 0)) if new_priority == "Y" else 0

    sig = get_signals_for_axis(lead_compare_df, chain, target_axis)
    if not sig.empty and "최적lag개월" in sig.columns:
        top_sig = sig.head(3).copy()
        rec_lag = int(round(top_sig["최적lag개월"].median()))
    else:
        top_sig = pd.DataFrame()
        rec_lag = None

    return {
        "old_raw": old_raw,
        "new_raw": new_raw,
        "old_rel": old_rel,
        "new_rel": new_rel,
        "old_alt": old_alt,
        "new_alt": new_alt,
        "old_priority": base_row["상대적_우선관리대상"],
        "new_priority": new_priority,
        "old_strength": base_row["우선관리강도"],
        "new_strength": new_strength,
        "q75": q75,
        "q25": q25,
        "new_q75_gap": new_q75_gap,
        "new_q25_gap": new_q25_gap,
        "scenario_desc": sc["desc"],
        "axis": target_axis,
        "signal_table": top_sig,
        "recommended_lag": rec_lag,
    }

@st.cache_data(show_spinner=False)
def load_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    raw = {}
    for s in xls.sheet_names:
        raw[s] = clean_columns(pd.read_excel(uploaded_file, sheet_name=s))

    def g(*names):
        for name in names:
            if name in raw:
                return raw[name]
        return pd.DataFrame()

    country = ensure_ym_column(g("COUNTRY_MONTHLY"))
    panel = ensure_ym_column(g("PANEL_MONTHLY"))
    alert = ensure_ym_column(g("ALERT_RESULT"))
    compare = clean_columns(g("체인별 비교표"))
    entropy = clean_columns(g("ENTROPY_WEIGHT"))
    norm_check = clean_columns(g("NOMALIZATION_CHECK"))
    norm_audit = clean_columns(g("NOMALIZATION_AUDIT"))
    method = clean_columns(g("METHOD_GUIDE"))
    signal_base = ensure_ym_column(g("SIGNAL_BASE"))
    signal_lag = ensure_ym_column(g("SIGNAL_LAG_TABLE"))
    lead_detail = clean_columns(g("LEAD_SIGNAL_LAG_DETAIL"))
    lead_compare = clean_columns(g("LEAD_SIGNAL_LAG_COMPARE"))
    hs_summary = ensure_ym_column(g("HS_MONTHLY_SUMMARY"))
    market = ensure_ym_column(g("MARKET_INDEX", "MAREKT_INDEX"))
    gscpi = ensure_ym_column(g("GSCPI_INDEX"))
    tpu = ensure_ym_column(g("TPU_INDEX"))

    for df in [country, panel, alert, compare, entropy, signal_base, signal_lag, lead_detail, lead_compare, hs_summary, market, gscpi, tpu]:
        if "체인구분" in df.columns:
            df["체인구분"] = df["체인구분"].astype(str).str.strip()

    numeric_candidates = set()
    all_frames = [country, panel, alert, compare, entropy, signal_base, signal_lag, lead_detail, lead_compare, hs_summary, market, gscpi, tpu, norm_check, norm_audit]
    for df in all_frames:
        for c in df.columns:
            if any(tok in str(c) for tok in ["점수", "비중", "의존도", "HHI", "금액", "물량", "가격", "INDEX", "Norm", "정규화", "lag", "q25", "q75", "median", "max", "min", "관측치", "개월수"]):
                numeric_candidates.add(c)

    for df in all_frames:
        for c in [x for x in numeric_candidates if x in df.columns]:
            df[c] = safe_numeric(df[c])

    if "보정사유" in alert.columns:
        alert["보정사유"] = alert["보정사유"].replace({"정책": "정책이벤트"})
    if "비고" in alert.columns:
        alert["비고"] = alert["비고"].astype(str).str.replace("정책 축", "정책이벤트 축", regex=False)

    alert = add_priority_metrics(alert, panel, compare)

    return {
        "sheets": raw,
        "country": country,
        "panel": panel,
        "alert": alert,
        "compare": compare,
        "entropy": entropy,
        "norm_check": norm_check,
        "norm_audit": norm_audit,
        "method": method,
        "signal_base": signal_base,
        "signal_lag": signal_lag,
        "lead_detail": lead_detail,
        "lead_compare": lead_compare,
        "hs_summary": hs_summary,
        "market": market,
        "gscpi": gscpi,
        "tpu": tpu,
    }

# ---------------------------------------------------------
# 앱 시작
# ---------------------------------------------------------
style_metric_container()

uploaded_file = st.sidebar.file_uploader("최종 확정본 엑셀 업로드", type=["xlsx"])

if uploaded_file is None:
    st.title("망보는사람들 공급망 리스크 대시보드")
    st.info("최종 확정본 엑셀 파일을 업로드하면 메뉴가 활성화됩니다.")
    st.stop()

data = load_excel(uploaded_file)

panel = data["panel"]
alert = data["alert"]
country = data["country"]
compare = data["compare"]
entropy = data["entropy"]
norm_check = data["norm_check"]
norm_audit = data["norm_audit"]
method = data["method"]
signal_base = data["signal_base"]
signal_lag = data["signal_lag"]
lead_detail = data["lead_detail"]
lead_compare = data["lead_compare"]
hs_summary = data["hs_summary"]
market = data["market"]
gscpi = data["gscpi"]
tpu = data["tpu"]

month_list = get_month_list(panel if not panel.empty else alert)
chain_list = get_chain_list(panel if not panel.empty else alert)

if not month_list or not chain_list:
    st.error("연월 또는 체인구분을 읽지 못했습니다. 최종 확정본 시트 구조를 확인해주세요.")
    st.stop()

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

needs_month = menu in [
    "1. 종합 상황판",
    "3. 국가/공급선 상세 분석",
    "4. 충격 원인 추적",
    "7. 대체국 추천 시스템",
    "9. 데이터 검증 / 방법론",
]
needs_chain = menu in [
    "3. 국가/공급선 상세 분석",
    "4. 충격 원인 추적",
    "5. 선행 신호 후보 탐지",
    "7. 대체국 추천 시스템",
    "9. 데이터 검증 / 방법론",
]

selected_month = st.sidebar.selectbox("기준 연월", month_list, index=len(month_list)-1) if needs_month else None
selected_chain = st.sidebar.selectbox("기준 체인", chain_list, index=0) if needs_chain else None

st.sidebar.caption("연월은 앱 전체에서 YYYY-MM 형식으로 통일해 표시합니다.")

panel_month = month_chain_slice(panel, selected_month)
alert_month = month_chain_slice(alert, selected_month)
panel_chain = month_chain_slice(panel, None, selected_chain).sort_values("연월") if selected_chain else pd.DataFrame()
alert_chain = month_chain_slice(alert, None, selected_chain).sort_values("연월") if selected_chain else pd.DataFrame()
compare_chain = compare[compare["체인구분"] == selected_chain].copy() if selected_chain else pd.DataFrame()

# ---------------------------------------------------------
# 1. 종합 상황판
# ---------------------------------------------------------
if menu == "1. 종합 상황판":
    st.header("1. 종합 상황판")

    info_box("핵심 해석 원칙", [
        "최종위험점수(원점수)는 가격·수급·물류·정책이벤트 4개 축을 체인별 엔트로피 가중치(EWM)로 결합한 결과입니다. 따라서 단순 평균이 아니라, 해당 체인에서 구분력이 큰 축이 더 크게 반영됩니다.",
        "상대위험지수는 최종위험점수(원점수)를 같은 체인 내부의 과거 분포 기준으로 0~100 범위로 정규화한 상대지표입니다. 0은 무위험이 아니라 체인 내부 상대적 저점, 100은 절대 최대위험이 아니라 체인 내부 상대적 고점을 의미합니다.",
        "우선관리대상 Y는 상대위험지수가 체인별 상위 사분위(Q75) 이상이면서 동시에 대체조달가능성 점수가 체인별 하위 사분위(Q25) 이하인 경우입니다. 즉, 위험 수준이 높고 동시에 대체조달 여건이 취약할 때만 Y입니다.",
        "Q75 초과폭 = 현재 상대위험지수 - 체인별 Q75, Q25 미달폭 = 체인별 대체조달가능성 Q25 - 현재 대체조달가능성 점수입니다. 관리강도는 우선관리대상 Y인 경우에만 max(Q75초과폭,0)+max(Q25미달폭,0)으로 계산합니다."
    ])

    rows = []
    for chain in chain_list:
        p = panel_month[panel_month["체인구분"] == chain]
        a = alert_month[alert_month["체인구분"] == chain]
        c = compare[compare["체인구분"] == chain]
        if p.empty or a.empty or c.empty:
            continue
        p = p.iloc[0]
        a = a.iloc[0]
        c = c.iloc[0]
        rows.append({
            "체인구분": chain,
            "최종위험점수(원점수)": p["최종위험점수(원점수)"],
            "상대위험지수(정규화값)": a["상대위험지수(정규화값)"],
            "최종경보등급": a["최종경보등급"],
            "상대적_우선관리대상": a["상대적_우선관리대상"],
            "대체조달가능성_점수": a["대체조달가능성_점수"],
            "대체조달가능성_Q25": a["대체조달가능성_Q25"],
            "Q75_초과폭": a["Q75_초과폭"],
            "Q25_미달폭": a["Q25_미달폭"],
            "우선관리강도": a["우선관리강도"],
            "보정사유": a["보정사유"],
            "상대위험지수_Q75": c["상대위험지수_Q75"],
            "가격리스크점수": p["가격리스크점수"],
            "수급리스크점수": p["수급리스크점수"],
            "물류리스크점수": p["물류리스크점수"],
            "정책이벤트리스크점수": p["정책이벤트리스크점수"],
        })
    summary = pd.DataFrame(rows)

    st.subheader(f"기준 연월: {selected_month}")
    st.markdown("**산식 요약**  \n최종위험점수(원점수) = 가격리스크×체인가중치 + 수급리스크×체인가중치 + 물류리스크×체인가중치 + 정책이벤트리스크×체인가중치")

    for chain in chain_list:
        row = summary[summary["체인구분"] == chain]
        if row.empty:
            continue
        r = row.iloc[0]
        st.markdown(f"### {chain}")

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        with m1:
            st.metric("최종위험점수(원점수)", fmt_num(r["최종위험점수(원점수)"]))
        with m2:
            st.metric("상대위험지수", fmt_num(r["상대위험지수(정규화값)"]), r["최종경보등급"])
        with m3:
            st.metric("우선관리대상", str(r["상대적_우선관리대상"]))
            st.caption(metric_help_text("우선관리대상", r))
        with m4:
            st.metric("관리강도", fmt_num(r["우선관리강도"]))
            st.caption(metric_help_text("관리강도", r))
        with m5:
            st.metric("Q75 초과폭", fmt_num(r["Q75_초과폭"]))
            st.caption(metric_help_text("Q75초과폭", r))
        with m6:
            st.metric("Q25 미달폭", fmt_num(r["Q25_미달폭"]))
            st.caption(metric_help_text("Q25미달폭", r))

        contrib = pd.DataFrame({
            "축": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"],
            "점수": [r["가격리스크점수"], r["수급리스크점수"], r["물류리스크점수"], r["정책이벤트리스크점수"]]
        })
        fig = px.bar(
            contrib, x="축", y="점수", color="축",
            color_discrete_map=AXIS_COLOR_MAP,
            title=f"{chain} · 최종위험점수 구성축"
        )
        fig.update_layout(showlegend=False, height=320)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 우선관리 사분면")
    fig2 = px.scatter(
        summary,
        x="대체조달가능성_점수",
        y="상대위험지수(정규화값)",
        color="체인구분",
        text="상대적_우선관리대상",
        size=np.maximum(summary["최종위험점수(원점수)"], 1),
        color_discrete_map=CHAIN_COLOR_MAP,
        hover_data=["Q75_초과폭", "Q25_미달폭", "우선관리강도", "보정사유"]
    )

    for _, r in summary.iterrows():
        fig2.add_hline(y=r["상대위험지수_Q75"], line_dash="dot", line_color=CHAIN_COLOR_MAP.get(r["체인구분"], "#999"))
        fig2.add_vline(x=r["대체조달가능성_Q25"], line_dash="dot", line_color=CHAIN_COLOR_MAP.get(r["체인구분"], "#999"))

    for ann in build_quadrant_annotations(summary):
        fig2.add_annotation(**ann)

    fig2.update_layout(
        height=430,
        xaxis_title="대체조달가능성 점수",
        yaxis_title="상대위험지수(정규화값)"
    )
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### 체인별 종합표")
    st.dataframe(
        summary[[
            "체인구분", "최종위험점수(원점수)", "상대위험지수(정규화값)", "최종경보등급",
            "상대적_우선관리대상", "보정사유", "Q75_초과폭", "Q25_미달폭", "우선관리강도"
        ]],
        use_container_width=True,
        hide_index=True
    )

# ---------------------------------------------------------
# 2. 체인별 심층 분석
# ---------------------------------------------------------
elif menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")

    info_box("핵심 해석 원칙", [
        "이 메뉴는 특정 연월의 단면이 아니라 전체 분석기간 동안 체인별 위험 구조가 어떻게 달랐는지를 비교하는 메뉴입니다. 따라서 기준 연월 선택 없이 체인 전체 추이와 분위수를 중심으로 해석합니다.",
        "최종위험점수(원점수)와 상대위험지수의 추이 방향이 유사하게 보이는 것은 정상입니다. 상대위험지수는 체인 내부에서 원점수를 0~100으로 선형 정규화한 값이기 때문에, 고점과 저점의 위치는 같고 스케일만 다릅니다.",
        "체인 간 비교에서는 상대위험지수만 보면 안 됩니다. 같은 30점이라도 체인별 원점수 분포와 Q25·Q50·Q75 기준선이 다르므로, 원점수 분위수와 함께 읽어야 실제 차이를 해석할 수 있습니다."
    ])

    st.subheader("체인별 기준선 비교")
    st.dataframe(compare, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        qdf = compare[["체인구분", "상대위험지수_Q25", "상대위험지수_Q50", "상대위험지수_Q75"]].melt(
            id_vars="체인구분", var_name="구간", value_name="값"
        )
        fig = px.bar(
            qdf, x="구간", y="값", color="체인구분",
            barmode="group", color_discrete_map=CHAIN_COLOR_MAP,
            title="체인별 상대위험지수 분위수"
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        qdf2 = compare[["체인구분", "최종위험점수(원점수)_Q25", "최종위험점수(원점수)_Q50", "최종위험점수(원점수)_Q75"]].melt(
            id_vars="체인구분", var_name="구간", value_name="값"
        )
        fig2 = px.bar(
            qdf2, x="구간", y="값", color="체인구분",
            barmode="group", color_discrete_map=CHAIN_COLOR_MAP,
            title="체인별 최종위험점수(원점수) 분위수"
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("체인별 월별 추이")
    tab1, tab2 = st.tabs(["최종위험점수(원점수) 추이", "상대위험지수 추이"])

    with tab1:
        fig3 = px.line(
            panel, x="연월", y="최종위험점수(원점수)", color="체인구분",
            markers=True, color_discrete_map=CHAIN_COLOR_MAP,
            title="체인별 최종위험점수(원점수) 추이"
        )
        st.plotly_chart(fig3, use_container_width=True)

    with tab2:
        fig4 = px.line(
            panel, x="연월", y="상대위험지수(정규화값)", color="체인구분",
            markers=True, color_discrete_map=CHAIN_COLOR_MAP,
            title="체인별 상대위험지수 추이"
        )
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("4대 리스크 평균 비교")
    axis_cols = ["평균_가격리스크점수", "평균_수급리스크점수", "평균_물류리스크점수", "평균_정책이벤트리스크점수"]
    axis_map_name = {
        "평균_가격리스크점수": "가격리스크점수",
        "평균_수급리스크점수": "수급리스크점수",
        "평균_물류리스크점수": "물류리스크점수",
        "평균_정책이벤트리스크점수": "정책이벤트리스크점수",
    }
    a = compare[["체인구분"] + axis_cols].melt(id_vars="체인구분", var_name="항목", value_name="값")
    a["항목"] = a["항목"].map(axis_map_name)
    fig5 = px.bar(
        a, x="항목", y="값", color="체인구분",
        barmode="group", color_discrete_map=CHAIN_COLOR_MAP,
        title="체인별 4대 리스크 평균 비교"
    )
    st.plotly_chart(fig5, use_container_width=True)

# ---------------------------------------------------------
# 3. 국가/공급선 상세 분석
# ---------------------------------------------------------
elif menu == "3. 국가/공급선 상세 분석":
    st.header("3. 국가/공급선 상세 분석")

    info_box("핵심 해석 원칙", [
        "이 메뉴는 선택한 체인과 연월에서 어떤 국가 공급선이 상대적으로 더 안정적이거나 더 취약한지를 비교하는 화면입니다.",
        "국가기본위험 평가등급은 기본평가점수의 수준을, 국가공급선 최종판정은 최종보정점수의 수준을 등급으로 표현한 것입니다. 즉, 전자는 기본위험, 후자는 공급구조 보정까지 반영된 종합 판정입니다.",
        "국가별 수입비중은 현재 조달이 어느 국가에 집중되어 있는지를 보여주고, 최종보정점수는 그 공급선의 상대적 부담 정도를 보여줍니다. 두 값을 함께 봐야 실제 관리 우선순위를 판단할 수 있습니다."
    ])

    ctry = month_chain_slice(country, selected_month, selected_chain)
    ctry_agg = aggregate_country_view(ctry)

    if ctry_agg.empty:
        st.warning("선택한 조건에 해당하는 국가 데이터가 없습니다.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            top_n = st.slider("상위 국가 표시 수", 5, min(30, len(ctry_agg)), min(15, len(ctry_agg)))
        with c2:
            fta_only = st.checkbox("FTA 체결국만 보기", value=False)
        with c3:
            supplier_only = st.checkbox("상위공급국만 보기", value=False)

        show = ctry_agg.copy()
        if fta_only and "FTA여부" in show.columns:
            show = show[show["FTA여부"] == "Y"].copy()
        if supplier_only and "상위공급국여부" in show.columns:
            show = show[show["상위공급국여부"] == "Y"].copy()

        show = show.head(top_n).copy()

        left, right = st.columns(2)

        with left:
            fig = px.bar(
                show.sort_values("국가별수입비중"),
                x="국가별수입비중",
                y="국가명",
                orientation="h",
                color="국가공급선_최종판정",
                color_discrete_map=GRADE_COLOR_MAP,
                title=f"{selected_month} {selected_chain} 국가별 수입비중"
            )
            st.plotly_chart(fig, use_container_width=True)

        with right:
            fig2 = px.scatter(
                show,
                x="최종보정점수",
                y="국가별수입비중",
                color="FTA여부" if "FTA여부" in show.columns else None,
                size=np.maximum(safe_numeric(show.get("국가수입금액", 1)).fillna(1), 1),
                hover_data=[c for c in ["국가명", "지역권", "국가기본위험_평가등급", "국가공급선_최종판정"] if c in show.columns],
                title="국가별 최종보정점수 vs 수입비중"
            )
            fig2.update_xaxes(title="최종보정점수(낮을수록 상대적으로 안정)")
            fig2.update_yaxes(title="국가별 수입비중")
            st.plotly_chart(fig2, use_container_width=True)

        out_cols = [c for c in [
            "국가명", "지역권", "FTA여부", "상위공급국여부", "국가별수입비중",
            "국가수입금액", "기본평가점수", "국가기본위험_평가등급",
            "최종보정점수", "국가공급선_최종판정", "비고"
        ] if c in show.columns]

        st.dataframe(show[out_cols], use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# 4. 충격 원인 추적
# ---------------------------------------------------------
elif menu == "4. 충격 원인 추적":
    st.header("4. 충격 원인 추적")

    info_box("핵심 해석 원칙", [
        "충격 원인은 먼저 4대 리스크 축의 절대수준을 보고, 그 다음 최종위험점수(원점수), 마지막으로 상대위험지수의 체인 내 위치를 해석하는 순서로 읽는 것이 가장 명확합니다.",
        "보정사유는 해당 월에 상대적으로 두드러진 축을 요약한 라벨입니다. 따라서 최종 판단은 가격·수급·물류·정책이벤트 4개 축 수치를 함께 보고 해야 합니다.",
        "아래의 Q25·Q50·Q75 값은 이 체인 안에서 어느 수준부터 중위권·상위권으로 볼 것인지 알려주는 경계값입니다. 예를 들어 Q75를 넘으면 그 체인 내부에서는 상위 위험구간으로 해석합니다."
    ])

    prow = panel_month[panel_month["체인구분"] == selected_chain]
    arow = alert_month[alert_month["체인구분"] == selected_chain]
    crow = compare_chain

    if prow.empty or arow.empty or crow.empty:
        st.warning("선택한 조건의 데이터가 없습니다.")
    else:
        prow = prow.iloc[0]
        arow = arow.iloc[0]
        crow = crow.iloc[0]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("최종위험점수(원점수)", fmt_num(prow["최종위험점수(원점수)"]))
        c2.metric("상대위험지수", fmt_num(arow["상대위험지수(정규화값)"]), arow["최종경보등급"])
        c3.metric("보정사유", str(arow["보정사유"]))
        c4.metric("우선관리대상", str(arow["상대적_우선관리대상"]))

        comp = pd.DataFrame({
            "축": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"],
            "값": [prow["가격리스크점수"], prow["수급리스크점수"], prow["물류리스크점수"], prow["정책이벤트리스크점수"]],
        })
        fig = px.bar(
            comp, x="축", y="값", color="축",
            color_discrete_map=AXIS_COLOR_MAP,
            title=f"{selected_month} {selected_chain} 4대 리스크"
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown(
            f"- 상대위험지수 분위수 경계: **Q25 ({fmt_num(crow['상대위험지수_Q25'])}) / Q50 ({fmt_num(crow['상대위험지수_Q50'])}) / Q75 ({fmt_num(crow['상대위험지수_Q75'])})**  \n"
            f"- 원점수 분위수 경계: **Q25 ({fmt_num(crow['최종위험점수(원점수)_Q25'])}) / Q50 ({fmt_num(crow['최종위험점수(원점수)_Q50'])}) / Q75 ({fmt_num(crow['최종위험점수(원점수)_Q75'])})**"
        )

        with st.expander("TPU_INDEX(정책이벤트리스크 원천지표) 해석 배경 보기"):
            st.write(get_tpu_story_text())

# ---------------------------------------------------------
# 5. 선행 신호 후보 탐지
# ---------------------------------------------------------
elif menu == "5. 선행 신호 후보 탐지":
    st.header("5. 선행 신호 후보 탐지")

    info_box("핵심 해석 원칙", [
        "이 메뉴는 연월 단면이 아니라 체인 전체 기간에서 어떤 지표가 몇 개월 선행해 위험 변화와 연결되는지를 보는 구조 분석 메뉴입니다. 따라서 기준 연월 선택은 두지 않습니다.",
        "최적 lag는 특정 지표가 움직인 뒤 몇 개월 후 최종위험점수와 가장 강하게 연결되었는지를 뜻합니다. 예를 들어 lag 2가 최적이면, 해당 지표는 약 2개월 뒤 위험 변화와 가장 관련이 컸다는 의미입니다.",
        "상관이 곧 인과를 뜻하는 것은 아니지만, 모니터링 우선순위와 조기경보 시점을 설계하는 데는 충분히 유용합니다. lag가 길면 사전경보용, lag가 짧으면 즉시대응용 지표로 활용할 수 있습니다."
    ])

    chain_compare = lead_compare[lead_compare["체인구분"] == selected_chain].copy()
    chain_detail = lead_detail[lead_detail["체인구분"] == selected_chain].copy()

    topn = st.slider("상위 후보 수", 5, min(20, len(chain_compare)) if len(chain_compare) > 0 else 5, 10)
    rank_df = chain_compare.sort_values("최적lag절대상관", ascending=False).head(topn).copy()

    if not rank_df.empty:
        rank_df["최적lag개월_문자"] = rank_df["최적lag개월"].astype(int).astype(str)

    c1, c2 = st.columns(2)

    with c1:
        fig = px.bar(
            rank_df.sort_values("최적lag절대상관"),
            x="최적lag절대상관",
            y="지표",
            orientation="h",
            color="최적lag개월_문자" if not rank_df.empty else None,
            color_discrete_map=LAG_COLOR_MAP,
            title=f"{selected_chain} 최적 lag 상위 후보"
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        options = rank_df["지표"].tolist() if not rank_df.empty else []
        sel_signal = st.selectbox("세부 시차를 볼 지표 선택", options if options else [""])
        if sel_signal and not chain_detail.empty:
            tmp = chain_detail[chain_detail["지표"] == sel_signal].sort_values("lag개월").copy()
            fig2 = px.line(
                tmp,
                x="lag개월",
                y="상관계수",
                markers=True,
                title=f"{selected_chain} · {sel_signal} lag별 상관"
            )
            st.plotly_chart(fig2, use_container_width=True)

    st.dataframe(rank_df, use_container_width=True, hide_index=True)

    st.subheader("활용 가이드")
    if not rank_df.empty:
        g = rank_df.iloc[0]
        axis_guess = get_signal_axis(g["지표"])
        st.markdown(
            f"""
- **가장 우선적으로 볼 지표**: `{g['지표']}`  
- **관련 축 추정**: `{axis_guess}`  
- **권장 선행 점검 시점**: 약 **{int(g['최적lag개월'])}개월 전**부터 집중 모니터링  
- **해석 포인트**: 이 지표는 최종위험점수와의 절대상관이 **{fmt_num(g['최적lag절대상관'], 3)}**로 가장 크게 나타났습니다.  
- **실무 활용 예시**:  
  1. 월간 조기경보 보고서에 해당 지표를 별도 선행지표로 포함  
  2. 최적 lag 개월수만큼 앞서 임계치 이탈 여부를 점검  
  3. 6번 메뉴의 대응 시뮬레이터에서 같은 축 대응전략과 연결해 사전 대응 시점을 설정
            """
        )
# ---------------------------------------------------------
# 6. 기업 대응 우선순위 추천 / 시뮬레이터
# ---------------------------------------------------------
elif menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.header("6. 기업 대응 우선순위 추천 / 시뮬레이터")

    info_box("핵심 해석 원칙", [
        "이 메뉴는 특정 한 달을 보는 화면이 아니라, 전체 분석기간 중 어떤 시점이 실제로 먼저 대응되어야 하는지를 우선순위로 정렬하는 메뉴입니다. 따라서 기준 연월 선택은 두지 않습니다.",
        "우선순위는 상대위험지수, 대체조달가능성, Q75 초과폭, Q25 미달폭을 함께 고려합니다. 같은 Y라도 위험 초과가 큰지, 조달 취약이 큰지에 따라 대응 순서는 달라질 수 있습니다.",
        "시뮬레이터는 기업이 실제로 통제 가능한 대응유형을 선택했을 때, 특정 축의 부담을 보수적으로 얼마나 완화할 수 있는지를 가정해 보는 기능입니다. 임의로 상대위험지수를 크게 깎는 방식이 아니라, 대응유형별 보수적 효과를 반영합니다.",
        "또한 선행신호 분석 결과와 연결해, 어떤 축은 몇 개월 전부터 대응을 시작하는 것이 더 현실적인지 함께 제안합니다."
    ])

    chain_filter = st.multiselect("체인 선택", chain_list, default=chain_list)
    only_y = st.checkbox("현재 우선관리대상 Y만 보기", value=False)

    view_df = alert[alert["체인구분"].isin(chain_filter)].copy()
    if only_y:
        view_df = view_df[view_df["상대적_우선관리대상"] == "Y"].copy()

    view_df["추천우선순위"] = (
        safe_numeric(view_df["우선관리강도"]).fillna(0) * 0.45 +
        safe_numeric(view_df["상대위험지수(정규화값)"]).fillna(0) * 0.35 +
        (100 - safe_numeric(view_df["대체조달가능성_점수"]).fillna(0)) * 0.20
    )

    view_df = view_df.sort_values(["추천우선순위", "연월"], ascending=[False, True]).copy()

    st.subheader("현재 우선순위 추천")
    st.dataframe(
        view_df[[
            "연월", "체인구분", "최종위험점수(원점수)", "상대위험지수(정규화값)",
            "대체조달가능성_점수", "Q75_초과폭", "Q25_미달폭",
            "우선관리강도", "상대적_우선관리대상", "보정사유", "추천우선순위"
        ]],
        use_container_width=True,
        hide_index=True
    )

    st.subheader("대응 시뮬레이터")

    if view_df.empty:
        st.warning("시뮬레이션 대상 데이터가 없습니다.")
    else:
        pick_idx = st.selectbox(
            "시뮬레이션 대상 행 선택",
            view_df.index.tolist(),
            format_func=lambda i: f"{view_df.loc[i, '연월']} | {view_df.loc[i, '체인구분']} | {view_df.loc[i, '보정사유']} | 우선순위 {fmt_num(view_df.loc[i, '추천우선순위'])}"
        )

        base = view_df.loc[pick_idx]
        base_panel = panel[
            (panel["연월"] == base["연월"]) &
            (panel["체인구분"] == base["체인구분"])
        ].copy()

        if base_panel.empty:
            st.warning("선택한 행의 PANEL_MONTHLY 데이터를 찾지 못했습니다.")
        else:
            base_panel = base_panel.iloc[0]

            recommended_actions = {
                "가격": [
                    "가격 계약구조 조정(장기계약·헤지·판가연동)",
                    "FTA 활용국 중심 대체조달 전환",
                ],
                "수급": [
                    "공급선 다변화(보조 공급국 확보·의존도 완화)",
                    "FTA 활용국 중심 대체조달 전환",
                ],
                "물류": [
                    "안전재고·물류경로 다변화",
                    "공급선 다변화(보조 공급국 확보·의존도 완화)",
                ],
                "정책이벤트": [
                    "규제모니터링·통관대응·대체시장 사전준비",
                    "FTA 활용국 중심 대체조달 전환",
                ],
            }

            dominant_axis_col = get_driver_axis(base_panel)
            dominant_axis_name = str(dominant_axis_col).replace("리스크점수", "") if dominant_axis_col else "수급"
            action_options = recommended_actions.get(dominant_axis_name, list(recommended_actions["수급"]))
            action_type = st.selectbox("가정할 대응유형", action_options)

            sim = simulate_intervention(base, base_panel, entropy, panel, lead_compare, action_type)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("현재 최종위험점수(원점수)", fmt_num(sim["old_raw"]))
            c2.metric("시뮬레이션 후 원점수", fmt_num(sim["new_raw"]), f"{fmt_num(sim['new_raw'] - sim['old_raw'])}")
            c3.metric("현재 상대위험지수", fmt_num(sim["old_rel"]))
            c4.metric("시뮬레이션 후 상대위험지수", fmt_num(sim["new_rel"]), f"{fmt_num(sim['new_rel'] - sim['old_rel'])}")

            c5, c6, c7, c8 = st.columns(4)
            c5.metric("현재 대체조달가능성", fmt_num(sim["old_alt"]))
            c6.metric("시뮬레이션 후 대체조달가능성", fmt_num(sim["new_alt"]), f"{fmt_num(sim['new_alt'] - sim['old_alt'])}")
            c7.metric("현재 우선관리대상", str(sim["old_priority"]))
            c8.metric("시뮬레이션 후 우선관리대상", str(sim["new_priority"]))

            c9, c10 = st.columns(2)
            c9.metric("현재 관리강도", fmt_num(sim["old_strength"]))
            c10.metric("시뮬레이션 후 관리강도", fmt_num(sim["new_strength"]))

            st.markdown(
                f"""
**선택된 대표 충격축**: `{dominant_axis_name}`  
**선택 대응유형 설명**: {sim['scenario_desc']}  
**판단 기준**: 시뮬레이션은 기업이 실제로 통제 가능한 대응유형을 기준으로, 해당 축의 부담을 보수적으로 일부 줄이는 가정을 적용했습니다.  
                """
            )

            if sim["recommended_lag"] is not None:
                st.markdown(
                    f"""
**선행신호 연계 제안**  
- 이 대응유형은 주로 **{sim['axis']} 축**과 연결됩니다.  
- 관련 선행지표의 대표 최적 lag를 보면, **약 {sim['recommended_lag']}개월 전부터 대응을 착수하는 것이 상대적으로 유리**합니다.  
- 즉, 같은 축의 선행지표가 경계수준에 근접하면 실제 조달·계약·물류 대응을 그 시점 이전부터 준비하는 방식으로 활용할 수 있습니다.
                    """
                )

            if sim["signal_table"] is not None and not sim["signal_table"].empty:
                st.markdown("##### 관련 선행신호 상위 후보")
                st.dataframe(
                    sim["signal_table"][["지표", "최적lag개월", "최적lag절대상관", "최적lag상관"]],
                    use_container_width=True,
                    hide_index=True
                )

# ---------------------------------------------------------
# 7. 대체국 추천 시스템
# ---------------------------------------------------------
elif menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")

    info_box("핵심 해석 원칙", [
        "이 메뉴의 목적은 어떤 국가가 상대적으로 덜 위험하고, 동시에 실제 대체조달 후보로 검토할 만한지를 찾는 것입니다.",
        "국가기본위험 평가등급은 보정 전 기본위험 수준을 뜻하고, 국가공급선 최종판정은 집중도·FTA 여부 등 조달구조 보정까지 반영한 공급선 관점의 종합 판정입니다.",
        "그래프에서 x축은 최종보정점수, y축은 국가별 수입비중입니다. 왼쪽 아래는 상대적으로 안정적이면서 현재 집중도가 낮은 후보군, 오른쪽 위는 위험도와 집중도가 동시에 큰 구간으로 해석할 수 있습니다."
    ])

    ctry = aggregate_country_view(month_chain_slice(country, selected_month, selected_chain))

    if ctry.empty:
        st.warning("선택한 조건에 해당하는 국가 데이터가 없습니다.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            min_share = st.slider("최소 수입비중", 0.0, 20.0, 0.5, 0.5)
        with c2:
            prefer_fta = st.checkbox("FTA 국가 우선", value=False)

        cand = ctry[ctry["국가별수입비중"] >= min_share].copy()
        if prefer_fta and "FTA여부" in cand.columns:
            cand = cand[cand["FTA여부"] == "Y"].copy()

        cand["추천점수"] = (
            (100 - safe_numeric(cand["최종보정점수"]).fillna(100)) * 0.50 +
            (100 - safe_numeric(cand["국가별수입비중"]).fillna(100)) * 0.20 +
            np.where(cand["FTA여부"] == "Y", 15, 0) +
            np.where(cand["상위공급국여부"] == "Y", -5, 5)
        )

        cand = cand.sort_values(["추천점수", "최종보정점수"], ascending=[False, True]).copy()

        left, right = st.columns(2)

        with left:
            fig = px.bar(
                cand.head(15).sort_values("추천점수"),
                x="추천점수",
                y="국가명",
                orientation="h",
                color="FTA여부" if "FTA여부" in cand.columns else None,
                title="대체국 추천 상위 후보"
            )
            st.plotly_chart(fig, use_container_width=True)

        with right:
            fig2 = px.scatter(
                cand.head(30),
                x="최종보정점수",
                y="국가별수입비중",
                color="FTA여부" if "FTA여부" in cand.columns else None,
                size=np.maximum(safe_numeric(cand.head(30).get("국가수입금액", 1)).fillna(1), 1),
                hover_data=[c for c in ["국가명", "국가기본위험_평가등급", "국가공급선_최종판정", "추천점수"] if c in cand.columns],
                title="대체국 후보 분포"
            )
            st.plotly_chart(fig2, use_container_width=True)

        st.dataframe(
            cand[[
                "국가명", "지역권", "FTA여부", "상위공급국여부", "국가별수입비중",
                "최종보정점수", "국가기본위험_평가등급", "국가공급선_최종판정", "추천점수"
            ]],
            use_container_width=True,
            hide_index=True
        )

# ---------------------------------------------------------
# 8. 원천데이터 탐색 / 다운로드
# ---------------------------------------------------------
elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")

    info_box("핵심 해석 원칙", [
        "이 메뉴는 앱에 표시되는 값이 어떤 시트에서 왔는지 직접 확인하는 화면입니다.",
        "시트를 선택하면 해당 시트가 무엇을 담고 있는지 설명과 함께 실제 데이터를 그대로 보여줍니다.",
        "MARKET_INDEX, GSCPI_INDEX, TPU_INDEX, HS_MONTHLY_SUMMARY 같은 시트도 연월을 YYYY-MM 형식으로 맞춰 보여주도록 처리했습니다."
    ])

    sheet_names = list(data["sheets"].keys())
    selected_sheet = st.selectbox("시트 선택", sheet_names)

    st.caption(SHEET_DESC.get(selected_sheet, "이 시트 설명은 아직 별도로 등록되지 않았습니다."))

    processed_sheet = get_processed_sheet_for_display(selected_sheet, data)
    if processed_sheet is not None and not processed_sheet.empty:
        sheet_df = processed_sheet.copy()
    else:
        sheet_df = ensure_ym_column(clean_columns(data["sheets"][selected_sheet].copy()))

    st.dataframe(sheet_df, use_container_width=True, hide_index=True)

    st.download_button(
        "현재 시트 CSV 다운로드",
        data=sheet_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{selected_sheet}.csv",
        mime="text/csv"
    )
    st.download_button(
        "전체 시트 Excel 다운로드",
        data=build_download_excel(data["sheets"]),
        file_name="battery_dashboard_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------------------------------------
# 9. 데이터 검증 / 방법론
# ---------------------------------------------------------
elif menu == "9. 데이터 검증 / 방법론":
    st.header("9. 데이터 검증 / 방법론")

    info_box("핵심 해석 원칙", [
        "이 메뉴는 값이 왜 이렇게 나왔는지를 확인하는 검증용 화면입니다. 체인별 가중치, 정규화 기준, 원점수와 상대위험지수의 연결관계를 한 번에 확인할 수 있습니다.",
        "상대위험지수는 절대위험 자체를 뜻하는 값이 아니라, 해당 체인 안에서 현재 원점수가 어느 위치에 있는지를 보여주는 상대적 위치지표입니다.",
        "따라서 실제 보고와 의사결정에서는 원점수, 상대위험지수, 그리고 Q25·Q50·Q75 경계값을 함께 보는 것이 가장 바람직합니다."
    ])

    tab1, tab2, tab3, tab4 = st.tabs(["방법론", "체인별 가중치", "정규화 검증", "선택 조건 감사"])

    with tab1:
        st.dataframe(method, use_container_width=True, hide_index=True)
        with st.expander("TPU_INDEX 서사배경 보기"):
            st.write(get_tpu_story_text())

    with tab2:
        stage_options = entropy["단계"].dropna().astype(str).unique().tolist() if not entropy.empty and "단계" in entropy.columns else []
        if stage_options:
            stage = st.selectbox("가중치 단계 선택", stage_options)
            w = entropy[entropy["단계"] == stage].copy()
            st.dataframe(w, use_container_width=True, hide_index=True)
            st.markdown("**가중치 합 점검**")
            st.dataframe(
                w.groupby(["단계", "체인구분"], as_index=False)["가중치"].sum(),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("가중치 단계 정보를 찾지 못했습니다.")

    with tab3:
        st.dataframe(norm_check, use_container_width=True, hide_index=True)
        if norm_audit is not None and not norm_audit.empty:
            st.markdown("**추가 정규화 감사 시트**")
            st.dataframe(norm_audit, use_container_width=True, hide_index=True)

    with tab4:
        prow = panel_month[panel_month["체인구분"] == selected_chain]
        arow = alert_month[alert_month["체인구분"] == selected_chain]
        crow = compare_chain

        if prow.empty or arow.empty or crow.empty:
            st.warning("선택한 조건의 감사 데이터가 없습니다.")
        else:
            prow = prow.iloc[0]
            arow = arow.iloc[0]
            crow = crow.iloc[0]

            chain_panel = panel[panel["체인구분"] == selected_chain].copy()
            raw_min = chain_panel["최종위험점수(원점수)"].min()
            raw_max = chain_panel["최종위험점수(원점수)"].max()
            raw_val = prow["최종위험점수(원점수)"]

            if pd.notna(raw_val) and pd.notna(raw_min) and pd.notna(raw_max) and raw_max != raw_min:
                calc_rel = (raw_val - raw_min) / (raw_max - raw_min) * 100
            else:
                calc_rel = np.nan

            audit_df = pd.DataFrame([
                ["연월", selected_month],
                ["체인구분", selected_chain],
                ["최종위험점수(원점수)", raw_val],
                ["체인내 원점수 최소값", raw_min],
                ["체인내 원점수 최대값", raw_max],
                ["원점수 기준 재계산 상대위험지수", calc_rel],
                ["시트상 상대위험지수", arow["상대위험지수(정규화값)"]],
                ["상대위험지수 Q25", crow["상대위험지수_Q25"]],
                ["상대위험지수 Q50", crow["상대위험지수_Q50"]],
                ["상대위험지수 Q75", crow["상대위험지수_Q75"]],
                ["대체조달가능성 Q25", arow["대체조달가능성_Q25"]],
                ["현재 대체조달가능성", arow["대체조달가능성_점수"]],
                ["현재 우선관리대상", arow["상대적_우선관리대상"]],
                ["현재 관리강도", arow["우선관리강도"]],
            ], columns=["항목", "값"])

            st.dataframe(audit_df, use_container_width=True, hide_index=True)
