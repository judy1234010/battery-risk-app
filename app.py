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
RISK_COLOR_MAP = {
    "낮음": "#4C78A8",
    "보통": "#F2C14E",
    "높음": "#F28E2B",
    "매우높음": "#E15759",
    "해석유보": "#9D9D9D"
}

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

def info_box(purpose, components, interpretation):
    html = f"""
    <div style="
        border: 1px solid rgba(255,255,255,0.18);
        border-radius: 10px;
        padding: 16px 18px;
        background: rgba(255,255,255,0.03);
        margin-bottom: 14px;
    ">
        <div style="font-weight:700; margin-bottom:10px;">분석 안내</div>
        <div style="margin-bottom:8px;"><b>분석 목적</b><br>{purpose}</div>
        <div style="margin-bottom:8px;"><b>핵심 구성</b><br>{components}</div>
        <div><b>해석 기준</b><br>{interpretation}</div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def quantile_or_current(df, col, q, current_val):
    if col not in df.columns:
        return current_val
    s = safe_numeric(df[col]).dropna()
    if s.empty:
        return current_val
    try:
        return float(s.quantile(q))
    except Exception:
        return current_val

def quantile_from_series_df(df, col, q, fallback=np.nan):
    if df is None or df.empty or col not in df.columns:
        return fallback
    s = safe_numeric(df[col]).dropna()
    if s.empty:
        return fallback
    try:
        return float(s.quantile(q))
    except Exception:
        return fallback

def get_entropy_weights(entropy_df):
    if entropy_df is None or entropy_df.empty or "변수명" not in entropy_df.columns or "가중치" not in entropy_df.columns:
        return {}
    return dict(zip(entropy_df["변수명"], safe_numeric(entropy_df["가중치"])))

def empirical_relative_risk_index(month_df, row_dict, entropy_df=None):
    if month_df is None or month_df.empty:
        return np.nan

    w = get_entropy_weights(entropy_df)
    axes = [
        ("가격리스크점수", w.get("가격리스크점수", 0.25)),
        ("수급리스크점수", w.get("수급리스크점수", 0.25)),
        ("물류리스크점수", w.get("물류리스크점수", 0.25)),
        ("정책이벤트리스크점수", w.get("정책이벤트리스크점수", 0.25)),
    ]

    score = 0.0
    used_weight = 0.0
    for col, weight in axes:
        if col not in month_df.columns:
            continue
        s = safe_numeric(month_df[col]).dropna()
        if s.empty:
            continue
        v = safe_numeric(pd.Series([row_dict.get(col, np.nan)])).iloc[0]
        if pd.isna(v):
            continue

        s_min = s.min()
        s_max = s.max()
        if pd.isna(s_min) or pd.isna(s_max):
            continue

        if abs(s_max - s_min) < 1e-12:
            norm = 50.0
        else:
            norm = (float(v) - float(s_min)) / (float(s_max) - float(s_min)) * 100.0
            norm = float(np.clip(norm, 0, 100))

        score += norm * weight
        used_weight += weight

    if used_weight == 0:
        return np.nan
    return score / used_weight

def scenario_rule_text(name):
    mapping = {
        "현 상태 유지": [
            "현재 공급구조와 운영 수준을 기준선으로 유지한다."
        ],
        "부분 다변화": [
            "상위 공급국 의존을 일부 완화하고 보조 공급선을 확대하는 시나리오다.",
            "현재값 기준으로 상위1국의존도 10% 완화, HHI 12% 완화, 수입국수 1개국 확대를 적용한다."
        ],
        "권역 분산": [
            "특정 국가와 권역에 집중된 공급구조를 완화하기 위해 조달 권역을 넓히는 시나리오다.",
            "현재값 기준으로 상위1국의존도 6% 완화, HHI 10% 완화, 수입국수 1개국 확대, 지역권수 1권역 확대를 적용한다."
        ],
        "물류 안정화": [
            "재고, 운송 루트, 리드타임 관리 등 운영 대응으로 물류 충격의 체감 영향을 일부 완화하는 시나리오다.",
            "현재값 기준으로 물류리스크점수 10% 완화를 적용하되, 해당 체인의 과거 분포를 고려해 과도한 개선은 제한한다."
        ],
        "종합 대응": [
            "부분 다변화, 권역 분산, 물류 안정화를 함께 반영한 종합 시나리오다.",
            "현재값 기준으로 상위1국의존도 12% 완화, HHI 15% 완화, 수입국수 2개국 확대, 지역권수 1권역 확대, 물류리스크 10% 완화를 함께 적용한다."
        ],
    }
    return mapping.get(name, [])

def build_scenario_parameter_guide():
    return pd.DataFrame({
        "시나리오": ["현 상태 유지", "부분 다변화", "권역 분산", "물류 안정화", "종합 대응"],
        "사전 적용 변화폭": [
            "변화 없음",
            "상위1국의존도 -10%, HHI -12%, 수입국수 +1",
            "상위1국의존도 -6%, HHI -10%, 수입국수 +1, 지역권수 +1",
            "물류리스크점수 -10%",
            "상위1국의존도 -12%, HHI -15%, 수입국수 +2, 지역권수 +1, 물류리스크점수 -10%"
        ],
        "제한 기준": [
            "기준선 유지",
            "구조지표는 동일 연월 전체 체인 분포 분위수 기준",
            "구조지표는 동일 연월 전체 체인 분포 분위수 기준",
            "물류지표는 해당 체인 시계열 분위수 기준",
            "구조지표는 동일 연월 분포, 물류지표는 체인 시계열 분포 기준"
        ],
        "제한 목적": [
            "기준선 비교",
            "과도한 개선 방지",
            "과도한 개선 방지",
            "과도한 개선 방지",
            "과도한 개선 방지",
        ]         
    })

def apply_logistics_buffer(current, chain_ts_df, reduction_rate=0.10):
    if pd.isna(current):
        return np.nan
    target_value = current * (1 - reduction_rate)
    floor_val = quantile_from_series_df(chain_ts_df, "물류리스크점수", 0.25, current)
    if pd.isna(floor_val):
        return target_value
    allowed_floor = floor_val if current > floor_val else current
    return max(target_value, allowed_floor)

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
    candidates = []

    top1 = safe_numeric(pd.Series([row.get("상위1국의존도", np.nan)])).iloc[0]
    hhi = safe_numeric(pd.Series([row.get("HHI", np.nan)])).iloc[0]
    fta = safe_numeric(pd.Series([row.get("fta_ratio", np.nan)])).iloc[0]
    regions = safe_numeric(pd.Series([row.get("지역권수", np.nan)])).iloc[0]

    if pd.notna(top1) and top1 >= 70:
        candidates.append({
            "score": 100,
            "권고 액션": "상위국 의존도 완화",
            "실행 아이디어": "상위 1개국 물량 일부를 2~3위국 또는 신규국으로 분산",
            "예상 개선 지표": "상위1국의존도, HHI 개선"
        })

    if pd.notna(hhi) and hhi >= 5000:
        candidates.append({
            "score": 90,
            "권고 액션": "공급선 집중 완화",
            "실행 아이디어": "단일·소수 공급국 구조를 다변화",
            "예상 개선 지표": "HHI, 수입국수 개선"
        })

    if pd.notna(regions) and regions <= 3:
        candidates.append({
            "score": 70,
            "권고 액션": "권역 다변화",
            "실행 아이디어": "동일 권역 집중 시 대체 권역 공급선 확보",
            "예상 개선 지표": "지역권수 개선"
        })

    if pd.notna(fta) and fta <= 40:
        candidates.append({
            "score": 60,
            "권고 액션": "FTA 활용 재점검",
            "실행 아이디어": "관세·원산지 조건을 검토해 FTA 활용 확대",
            "예상 개선 지표": "fta_ratio 개선"
        })

    if not candidates:
        candidates.append({
            "score": 10,
            "권고 액션": "정기 모니터링",
            "실행 아이디어": "월별 리스크 추이와 공급국 변동을 점검",
            "예상 개선 지표": "현 수준 유지 관리"
        })

    roadmap = pd.DataFrame(candidates).sort_values("score", ascending=False).reset_index(drop=True)
    roadmap["우선순위"] = [f"{i}순위" for i in range(1, len(roadmap) + 1)]
    return roadmap[["우선순위", "권고 액션", "실행 아이디어", "예상 개선 지표"]]

def _safe_get_num(row_dict, key):
    v = safe_numeric(pd.Series([row_dict.get(key, np.nan)])).iloc[0]
    return np.nan if pd.isna(v) else float(v)

def _clamp_lower_better(current, target, month_df, col, q_floor=0.25):
    if pd.isna(current):
        return np.nan
    floor_val = quantile_or_current(month_df, col, q_floor, current)
    allowed_floor = floor_val if current > floor_val else current
    candidate = min(current, target)
    return max(candidate, allowed_floor)

def _clamp_higher_better(current, target, month_df, col, q_cap=0.75):
    if pd.isna(current):
        return np.nan
    cap_val = quantile_or_current(month_df, col, q_cap, current)
    allowed_cap = cap_val if current < cap_val else current
    candidate = max(current, target)
    return min(candidate, allowed_cap)

def _minmax_risk(value, series, higher_is_risk=True):
    s = safe_numeric(series).dropna()
    if s.empty or pd.isna(value):
        return np.nan
    s_min = float(s.min())
    s_max = float(s.max())
    if abs(s_max - s_min) < 1e-12:
        return 50.0
    if higher_is_risk:
        score = (float(value) - s_min) / (s_max - s_min) * 100.0
    else:
        score = (s_max - float(value)) / (s_max - s_min) * 100.0
    return float(np.clip(score, 0, 100))

def calculate_supply_structure_risk(row_dict, month_df):
    """
    수급리스크 비교용 재평가 함수.
    동일 연월의 PANEL_MONTHLY 체인 분포에서
    상위1국의존도, HHI, 수입국수, 지역권수의 상대 위치를 0~100 위험도로 환산한 뒤
    동일가중 평균한다.
    """
    if month_df is None or month_df.empty:
        return np.nan

    parts = []
    if "상위1국의존도" in month_df.columns:
        parts.append(_minmax_risk(_safe_get_num(row_dict, "상위1국의존도"), month_df["상위1국의존도"], True))
    if "HHI" in month_df.columns:
        parts.append(_minmax_risk(_safe_get_num(row_dict, "HHI"), month_df["HHI"], True))
    if "수입국수" in month_df.columns:
        parts.append(_minmax_risk(_safe_get_num(row_dict, "수입국수"), month_df["수입국수"], False))
    if "지역권수" in month_df.columns:
        parts.append(_minmax_risk(_safe_get_num(row_dict, "지역권수"), month_df["지역권수"], False))

    parts = [p for p in parts if pd.notna(p)]
    if not parts:
        return np.nan
    return float(np.mean(parts))

def calculate_simulation_risk_index(row_dict, month_df, entropy_df=None):
    """
    6번 메뉴 전용 종합위험지수(참고용).
    4대 리스크 체계는 유지하되, 수급리스크는 구조지표 변화에 따른 비교용 재평가값을 사용한다.
    """
    w = get_entropy_weights(entropy_df)
    weight_price = float(w.get("가격리스크점수", 0.25))
    weight_supply = float(w.get("수급리스크점수", 0.25))
    weight_logi = float(w.get("물류리스크점수", 0.25))
    weight_policy = float(w.get("정책이벤트리스크점수", 0.25))

    price = _safe_get_num(row_dict, "가격리스크점수")
    supply = _safe_get_num(row_dict, "수급리스크점수")
    logistics = _safe_get_num(row_dict, "물류리스크점수")
    policy = _safe_get_num(row_dict, "정책이벤트리스크점수")

    values = [
        (price, weight_price),
        (supply, weight_supply),
        (logistics, weight_logi),
        (policy, weight_policy)
    ]
    num = sum(v * wgt for v, wgt in values if pd.notna(v))
    den = sum(wgt for v, wgt in values if pd.notna(v))
    if den == 0:
        return np.nan
    return float(num / den)

def apply_structure_scenario(row_dict, month_df, chain_ts_df, scenario_name):
    """
    기업 대응 시나리오 적용.
    - 직접 변경: 상위1국의존도, HHI, 수입국수, 지역권수, 물류 운영 완충
    - 직접 미변경: 가격리스크점수, 정책이벤트리스크점수
    - 수급리스크점수는 구조지표 변화에 따라 비교용으로 재평가
    """
    row = dict(row_dict)

    cur_top1 = _safe_get_num(row, "상위1국의존도")
    cur_hhi = _safe_get_num(row, "HHI")
    cur_countries = _safe_get_num(row, "수입국수")
    cur_regions = _safe_get_num(row, "지역권수")
    cur_logistics = _safe_get_num(row, "물류리스크점수")
    cur_supply = _safe_get_num(row, "수급리스크점수")

    if month_df is None or month_df.empty:
        return row

    if scenario_name == "현 상태 유지":
        row["수급리스크점수"] = calculate_supply_structure_risk(row, month_df) if pd.isna(cur_supply) else cur_supply
        return row

    if scenario_name == "부분 다변화":
        if pd.notna(cur_top1):
            row["상위1국의존도"] = _clamp_lower_better(cur_top1, cur_top1 * 0.90, month_df, "상위1국의존도", 0.25)
        if pd.notna(cur_hhi):
            row["HHI"] = _clamp_lower_better(cur_hhi, cur_hhi * 0.88, month_df, "HHI", 0.25)
        if pd.notna(cur_countries):
            row["수입국수"] = _clamp_higher_better(cur_countries, cur_countries + 1, month_df, "수입국수", 0.75)

    elif scenario_name == "권역 분산":
        if pd.notna(cur_top1):
            row["상위1국의존도"] = _clamp_lower_better(cur_top1, cur_top1 * 0.94, month_df, "상위1국의존도", 0.25)
        if pd.notna(cur_hhi):
            row["HHI"] = _clamp_lower_better(cur_hhi, cur_hhi * 0.90, month_df, "HHI", 0.25)
        if pd.notna(cur_countries):
            row["수입국수"] = _clamp_higher_better(cur_countries, cur_countries + 1, month_df, "수입국수", 0.75)
        if pd.notna(cur_regions):
            row["지역권수"] = _clamp_higher_better(cur_regions, cur_regions + 1, month_df, "지역권수", 0.75)

    elif scenario_name == "물류 안정화":
        if pd.notna(cur_logistics):
            row["물류리스크점수"] = apply_logistics_buffer(cur_logistics, chain_ts_df, reduction_rate=0.10)

    elif scenario_name == "종합 대응":
        if pd.notna(cur_top1):
            row["상위1국의존도"] = _clamp_lower_better(cur_top1, cur_top1 * 0.88, month_df, "상위1국의존도", 0.25)
        if pd.notna(cur_hhi):
            row["HHI"] = _clamp_lower_better(cur_hhi, cur_hhi * 0.85, month_df, "HHI", 0.25)
        if pd.notna(cur_countries):
            row["수입국수"] = _clamp_higher_better(cur_countries, cur_countries + 2, month_df, "수입국수", 0.75)
        if pd.notna(cur_regions):
            row["지역권수"] = _clamp_higher_better(cur_regions, cur_regions + 1, month_df, "지역권수", 0.75)
        if pd.notna(cur_logistics):
            row["물류리스크점수"] = apply_logistics_buffer(cur_logistics, chain_ts_df, reduction_rate=0.10)

    row["수급리스크점수"] = calculate_supply_structure_risk(row, month_df)

    if "가격리스크점수" in row_dict:
        row["가격리스크점수"] = row_dict.get("가격리스크점수", np.nan)
    if "정책이벤트리스크점수" in row_dict:
        row["정책이벤트리스크점수"] = row_dict.get("정책이벤트리스크점수", np.nan)

    return row

def build_simulation_comparison_table(current_row, scenario_row):
    compare_items = [
        "상위1국의존도", "HHI", "수입국수", "지역권수",
        "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"
    ]
    preferred_direction = {
        "상위1국의존도": "감소가 바람직",
        "HHI": "감소가 바람직",
        "수입국수": "증가가 바람직",
        "지역권수": "증가가 바람직",
        "수급리스크점수": "감소가 바람직",
        "물류리스크점수": "감소가 바람직",
        "정책이벤트리스크점수": "직접 변경 없음"
    }
    rows = []
    for item in compare_items:
        cur_val = _safe_get_num(current_row, item)
        new_val = _safe_get_num(scenario_row, item)
        diff = np.nan
        pct = np.nan
        if pd.notna(cur_val) and pd.notna(new_val):
            diff = new_val - cur_val
            if abs(cur_val) < 1e-12:
                pct = 0.0 if abs(new_val) < 1e-12 else np.nan
            else:
                pct = (new_val - cur_val) / abs(cur_val) * 100.0
        rows.append({
            "항목": item,
            "현재": cur_val,
            "시나리오 적용 후": new_val,
            "변화폭": diff,
            "변화율(%)": pct,
            "개선방향": preferred_direction.get(item, "-")
        })
    return pd.DataFrame(rows)

def risk_grade_order(v):
    mapping = {"낮음": 1, "보통": 2, "높음": 3, "매우높음": 4, "해석유보": 0}
    return mapping.get(str(v).strip(), -1)

def aggregate_country_recommendation(df):
    if df is None or df.empty:
        return pd.DataFrame()

    work = df.copy()
    group_key = "국가코드" if "국가코드" in work.columns else "국가라벨"

    if "국가라벨" not in work.columns:
        if "국가명" in work.columns and "국가코드" in work.columns:
            work["국가라벨"] = work["국가명"].astype(str) + " (" + work["국가코드"].astype(str) + ")"
        elif "국가명" in work.columns:
            work["국가라벨"] = work["국가명"].astype(str)
        else:
            work["국가라벨"] = work.index.astype(str)

    if "금액비중" in work.columns:
        weight_col = "금액비중"
    elif "국가수입금액" in work.columns:
        weight_col = "국가수입금액"
    else:
        weight_col = None

    num_cols = [c for c in ["국가수입금액", "금액비중", "기본평가점수", "최종보정점수", "대체국 우선검토지수"] if c in work.columns]
    cat_cols = [c for c in ["국가명", "지역권", "FTA여부", "상위공급국여부"] if c in work.columns]

    out_rows = []
    for key, sub in work.groupby(group_key, dropna=False):
        row = {}
        row[group_key] = key

        if "국가명" in sub.columns and "국가코드" in sub.columns:
            row["국가라벨"] = f"{sub['국가명'].iloc[0]} ({sub['국가코드'].iloc[0]})"
        else:
            row["국가라벨"] = sub["국가라벨"].iloc[0]

        if weight_col is not None:
            weights = safe_numeric(sub[weight_col]).fillna(0)
            if float(weights.sum()) <= 0:
                weights = pd.Series(np.ones(len(sub)), index=sub.index)
        else:
            weights = pd.Series(np.ones(len(sub)), index=sub.index)

        for col in num_cols:
            vals = safe_numeric(sub[col])
            mask = vals.notna() & weights.notna()
            if mask.sum() == 0:
                row[col] = np.nan
            else:
                row[col] = float(np.average(vals[mask], weights=weights[mask]))

        for col in cat_cols:
            mode = sub[col].mode(dropna=True)
            row[col] = mode.iloc[0] if not mode.empty else sub[col].iloc[0]

        if "최종판정" in sub.columns:
            sub2 = sub.copy()
            sub2["_risk_order"] = sub2["최종판정"].apply(risk_grade_order)
            sub2 = sub2.sort_values("_risk_order", ascending=False)
            row["최종판정(위험도분류)"] = sub2["최종판정"].iloc[0]

        out_rows.append(row)

    graph_df = pd.DataFrame(out_rows)
    if not graph_df.empty and "국가라벨" in graph_df.columns:
        graph_df = graph_df.drop_duplicates(subset=["국가라벨"]).copy()
    if "대체국 우선검토지수" in graph_df.columns:
        graph_df = graph_df.sort_values("대체국 우선검토지수", ascending=False).reset_index(drop=True)
    return graph_df

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

# =========================================================
# 앱 시작
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
    info_box(
        "선택한 기준 연월에서 배터리 공급망의 전반적 위험 수준을 체인별로 비교하기 위한 화면이다.",
        "1. 체인별 최종위험점수 및 경보등급 비교<br>2. 4대 리스크 축의 상대적 수준 비교<br>3. 체인별 비교표를 통한 평균 위험도와 우선관리 현황 요약",
        "전체 체인 중 상대적으로 위험 수준이 높은 대상을 먼저 식별한 뒤, 이후 심층 분석이 필요한 체인을 선별하는 출발점으로 사용한다."
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
    st.caption("물류리스크점수와 정책이벤트리스크점수는 동일 연월 기준 공통 외생지표(GSCPI, TPU_INDEX)를 사용하므로, 같은 연월에서는 체인 간 값이 동일하게 나타날 수 있다.")

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
    info_box(
        "특정 체인의 위험 수준이 시계열상 어떻게 변화해 왔는지와 구조적 취약성이 무엇인지를 분석하기 위한 화면이다.",
        "1. 월별 최종위험점수 및 4대 리스크 축 추이<br>2. 수입구조와 집중도 지표의 장기 흐름 확인<br>3. 선택 연월 기준 핵심 구조지표 점검",
        "단기 급등 여부보다, 특정 체인이 장기간 어떤 구조적 위험을 내포하고 있는지 판단하는 데 초점을 둔다."
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
    info_box(
        "특정 체인과 연월에서 어떤 국가가 공급망 위험에 실질적으로 영향을 주고 있는지 식별하기 위한 화면이다.",
        "1. 국가별 수입비중 비교<br>2. 국가별 최종보정점수 비교<br>3. 수입비중과 최종보정점수의 결합 분석",
        "수입비중이 높고 최종보정점수도 높은 국가는 우선 점검 대상이며, 비중은 낮지만 점수가 높은 국가는 잠재 위험국으로 해석할 수 있다."
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

    if "국가명" in sub.columns and "국가코드" in sub.columns:
        sub["국가라벨"] = sub["국가명"].astype(str) + " (" + sub["국가코드"].astype(str) + ")"
    elif "국가명" in sub.columns:
        sub["국가라벨"] = sub["국가명"].astype(str)
    else:
        sub["국가라벨"] = sub.index.astype(str)

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
    if {"국가라벨", "금액비중"}.issubset(top.columns):
        fig = px.bar(
            top.sort_values("금액비중", ascending=False),
            x="국가라벨",
            y="금액비중",
            color="최종판정" if "최종판정" in top.columns else None,
            color_discrete_map=RISK_COLOR_MAP,
            template="plotly_dark",
            hover_data=[c for c in ["FTA여부", "상위공급국여부", "최종보정점수", "국가수입금액"] if c in top.columns]
        )
        fig.update_layout(height=420, legend_title_text="최종판정", xaxis_title="국가", yaxis_title="금액비중")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### ② 국가별 최종보정점수")
    st.caption("국가별 조달 리스크 수준을 직접 비교하기 위한 그래프다. 최종보정점수는 국가기본위험에 공급집중, 지역집중, 상위공급국 여부, FTA 여부 등의 보정요인을 반영한 값이다.")
    if {"국가라벨", "최종보정점수", "FTA여부"}.issubset(top.columns):
        fig_score = px.bar(
            top.sort_values("최종보정점수", ascending=False),
            x="국가라벨",
            y="최종보정점수",
            color="FTA여부",
            color_discrete_map=FTA_COLOR_MAP,
            template="plotly_dark",
            hover_data=[c for c in ["금액비중", "상위공급국여부", "최종판정", "국가수입금액"] if c in top.columns]
        )
        fig_score.update_layout(height=420, legend_title_text="FTA 여부", xaxis_title="국가", yaxis_title="최종보정점수")
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
    info_box(
        "특정 월의 위험 상승이 어떤 요인에서 비롯되었는지를 직접 해석하기 위한 화면이다.",
        "1. 가격, 수급, 물류, 정책이벤트의 4개 리스크 축 비교<br>2. 선택 연월의 핵심 구조지표와 원천지표 확인<br>3. GSCPI, TPU_INDEX 등 외생 변수의 동시 점검",
        "체인별 심층 분석이 장기 구조를 파악하는 데 초점을 둔다면, 본 화면은 특정 시점의 경보 원인을 분해하여 설명하는 데 목적이 있다."
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
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan), 2), row.get("최종경보등급", grade_final_risk(row.get("최종위험점수", np.nan))))
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
    info_box(
        "최종위험점수에 선행하여 움직일 가능성이 있는 원천 지표를 식별하기 위한 화면이다.",
        "① 후보 변수 범위<br>② 변수별 최적 lag 요약<br>③ lag 1~6 상세 비교<br>④ 분석 베이스 데이터 확인",
        "절대상관이 크고 일관된 방향성을 보이는 지표일수록 선행 모니터링 가치가 높다고 해석한다."
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
    info_box(
        "기업이 현실적으로 조정 가능한 공급구조 및 운영 항목을 기준으로 대응 우선순위와 시나리오별 위험 변화를 비교하기 위한 화면이다.",
        "1. 구조 취약성 진단 및 우선 개선 과제 제시<br>2. 상위1국의존도, HHI, 수입국수, 지역권수, 물류 운영 안정성 중심 시나리오 비교<br>3. 시나리오 전후 종합위험 변화 확인",
        "가격리스크와 정책이벤트리스크는 외생 변수 성격이 강해 직접 조정하지 않는 것으로 두고, 기업이 통제 가능한 구조·운영 항목 변화가 수급 및 종합위험에 미치는 영향을 비교한다."
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

    month_df = panel[panel["연월"] == month].copy()
    chain_ts_df = panel[panel["체인구분"] == chain].sort_values("연월").copy()

    base_supply_sim = calculate_supply_structure_risk(row, month_df)
    if pd.notna(base_supply_sim):
        row["수급리스크점수_원본"] = row.get("수급리스크점수", np.nan)
        row["수급리스크점수"] = base_supply_sim

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("최종위험점수", fmt_num(row.get("최종위험점수", np.nan), 2), row.get("최종경보등급", "-"))
    c2.metric("대체조달가능성", fmt_num(row.get("대체조달가능성_점수", np.nan), 2), row.get("대체조달가능성", "-"))
    c3.metric("보정사유", row.get("보정사유", "-"))
    c4.metric("우선관리대상", row.get("상대적_우선관리대상", "-"))

    st.markdown("#### 경영진 요약 코멘트")
    st.write(build_executive_comment(row))

    st.markdown("#### 구조 진단")
    diag = pd.DataFrame({
        "항목": ["상위1국의존도", "HHI", "FTA 비중", "지역권수", "수입국수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"],
        "현재값": [
            row.get("상위1국의존도", np.nan),
            row.get("HHI", np.nan),
            row.get("fta_ratio", np.nan),
            row.get("지역권수", np.nan),
            row.get("수입국수", np.nan),
            row.get("수급리스크점수", np.nan),
            row.get("물류리스크점수", np.nan),
            row.get("정책이벤트리스크점수", np.nan)
        ]
    })
    st.dataframe(diag, use_container_width=True)

    st.markdown("#### 권고 액션 로드맵")
    st.caption("우선순위는 구조적 취약성이 큰 항목부터 실무 검토 순서대로 재정렬한 결과이다.")
    roadmap_df = build_action_roadmap(row)
    st.dataframe(roadmap_df, use_container_width=True)

    st.markdown("#### 구조 개선 시뮬레이터")
    st.caption("본 시뮬레이터는 기업이 실제로 조정 가능한 구조지표와 운영 대응만 반영한다. 가격리스크점수와 정책이벤트리스크점수는 직접 변경하지 않으며, 구조 변화에 따라 수급리스크를 비교용으로 재평가한다.")
    st.caption("시나리오별 변화폭은 현재값 기준의 사전 설정 비율·증감폭을 적용한 뒤, 실제 관측 데이터 분포를 벗어나지 않도록 분위수 기준으로 한 번 더 제한한다.")
    st.caption("상위1국의존도·HHI·수입국수·지역권수는 동일 연월 전체 체인 분포를 기준으로, 물류리스크점수는 해당 체인의 월별 시계열 분포를 기준으로 제한한다.")
    st.caption("따라서 일부 지표는 이미 분포상 양호한 수준에 있으면 시나리오를 선택해도 추가 완화가 반영되지 않을 수 있다.")

    scenario = st.selectbox(
        "대표 시나리오 선택",
        ["현 상태 유지", "부분 다변화", "권역 분산", "물류 안정화", "종합 대응"]
    )

    with st.expander("시나리오별 적용 기준 보기", expanded=False):
        st.dataframe(build_scenario_parameter_guide(), use_container_width=True, hide_index=True)

    st.caption("시나리오 정의")
    for idx, line in enumerate(scenario_rule_text(scenario), start=1):
        st.write(f"{idx}. {line}")

    st.caption("종합위험지수 계산식(참고용): 가격리스크점수, 수급리스크점수(구조지표 기반 비교용 재평가), 물류리스크점수, 정책이벤트리스크점수를 엔트로피 가중치로 가중평균한다.")

    current_sim_idx = calculate_simulation_risk_index(row, month_df, entropy)
    scenario_row = apply_structure_scenario(row, month_df, chain_ts_df, scenario)
    scenario_sim_idx = calculate_simulation_risk_index(scenario_row, month_df, entropy)

    impact_df = build_simulation_comparison_table(row, scenario_row)
    st.dataframe(impact_df, use_container_width=True)

    s1, s2 = st.columns(2)
    s1.metric("현재 종합위험지수(참고용)", fmt_num(current_sim_idx, 2))
    delta_txt = "-"
    if pd.notna(current_sim_idx) and pd.notna(scenario_sim_idx):
        delta_txt = fmt_num(scenario_sim_idx - current_sim_idx, 2)
    s2.metric("시나리오 적용 후 종합위험지수(참고용)", fmt_num(scenario_sim_idx, 2), delta=delta_txt)

    st.caption("변화율(%) 계산식: (시나리오 적용 후 - 현재) / 현재 × 100. 상위1국의존도·HHI·수급리스크·물류리스크는 감소가 바람직하고, 수입국수·지역권수는 증가가 바람직하다. 정책이벤트리스크점수는 기업이 직접 조정하지 않으므로 원칙적으로 변화하지 않는다.")

    rate_df = impact_df.copy()
    rate_df["변화율(%)"] = safe_numeric(rate_df["변화율(%)"])
    rate_df = rate_df[rate_df["변화율(%)"].notna()].copy()

    st.markdown("#### 시나리오 적용 전후 변화율 그래프")
    st.caption("항목별 절대 규모가 서로 다르므로, 시나리오 효과의 비교 가능성을 높이기 위해 변화율(%) 기준으로 시각화하였다.")
    if not rate_df.empty:
        fig_rate = px.bar(
            rate_df,
            x="항목",
            y="변화율(%)",
            color="개선방향",
            template="plotly_dark",
            hover_data=["현재", "시나리오 적용 후", "변화폭"]
        )
        fig_rate.update_layout(height=430, legend_title_text="해석")
        st.plotly_chart(fig_rate, use_container_width=True)
    else:
        st.info("변화율을 계산할 수 있는 항목이 없습니다.")

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
elif menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")
    info_box(
        "기존 공급구조를 보완하거나 대체할 수 있는 공급국 후보를 비교하기 위한 화면이다.",
        "1. 국가별 대체국 우선검토지수 비교<br>2. 국가별 최종판정(위험도분류) 동시 확인<br>3. FTA 여부, 상위공급국 여부, 최종보정점수 등 보조 정보 제공",
        "막대 높이는 대체국 우선검토지수, 막대 색은 해당 국가의 최종판정(위험도분류)을 의미한다. 국가 단위 비교의 일관성을 위해 국가코드 기준 집계값으로 제시한다."
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

    if "국가명" in cur.columns and "국가코드" in cur.columns:
        cur["국가라벨"] = cur["국가명"].astype(str) + " (" + cur["국가코드"].astype(str) + ")"
    elif "국가명" in cur.columns:
        cur["국가라벨"] = cur["국가명"].astype(str)
    else:
        cur["국가라벨"] = cur.index.astype(str)

    fta_only = st.checkbox("FTA 체결국만 보기", value=False)
    exclude_top = st.checkbox(
        "상위공급국 제외",
        value=False,
        help="기존 주력 공급국이 아니라 신규 보완 후보를 찾고 싶을 때 사용한다."
    )

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

    cur["대체국 우선검토지수"] = (
        cur["안정성점수"].fillna(0)
        + cur["FTA가점"]
        + cur["비상위가점"]
        + cur["규모가점"]
    )

    graph_df = aggregate_country_recommendation(cur)

    show_cols = [c for c in ["국가코드", "국가라벨", "국가명", "지역권", "FTA여부", "상위공급국여부", "국가수입금액", "금액비중", "기본평가점수", "최종보정점수", "최종판정(위험도분류)", "대체국 우선검토지수"] if c in graph_df.columns]
    st.dataframe(graph_df[show_cols], use_container_width=True)

    top10 = graph_df.head(10).copy()
    if not top10.empty and {"국가라벨", "대체국 우선검토지수"}.issubset(top10.columns):
        fig = px.bar(
            top10.sort_values("대체국 우선검토지수", ascending=False),
            x="국가라벨",
            y="대체국 우선검토지수",
            color="최종판정(위험도분류)" if "최종판정(위험도분류)" in top10.columns else None,
            color_discrete_map=RISK_COLOR_MAP,
            category_orders={"최종판정(위험도분류)": ["낮음", "보통", "높음", "매우높음", "해석유보"]},
            template="plotly_dark",
            hover_data=[c for c in ["FTA여부", "상위공급국여부", "최종보정점수", "금액비중"] if c in top10.columns]
        )
        fig.update_layout(
            height=420,
            legend_title_text="최종판정(위험도분류)",
            xaxis_title="국가",
            yaxis_title="대체국 우선검토지수"
        )
        st.plotly_chart(fig, use_container_width=True)

    st.caption("대체국 우선검토지수는 국가 자체의 위험도만이 아니라 안정성, FTA 여부, 기존 공급집중 완화 가능성, 거래규모를 함께 반영한 참고용 상대지표이다.")
    st.caption("동일 국가가 복수 행으로 존재하는 경우에는 국가 단위 비교가 가능하도록 국가코드 기준으로 집계하여 표시한다.")
    st.caption("최종판정(위험도분류) 의미: 낮음 = 현재 기준에서 조달위험이 비교적 낮은 상태 / 보통 = 즉시 경보 수준은 아니지만 일부 위험요인을 지속 점검할 필요가 있는 상태 / 높음 = 공급망 운영상 주의가 필요한 수준으로 대체·분산 여부를 함께 검토할 필요가 있는 상태")

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")
    info_box(
        "대시보드에 연결된 실제 원천 시트와 데이터를 직접 점검하기 위한 화면이다.",
        "1. 시트별 원본 테이블 확인<br>2. 행·열 규모 확인<br>3. CSV 및 Excel 다운로드 제공",
        "결과 검증, 후속 분석, 제출 전 교차점검을 위한 기초 데이터 확인 단계로 사용한다."
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
    info_box(
        "데이터 연결과 계산 구조가 정상적으로 작동하는지 점검하기 위한 화면이다.",
        "1. 필수 시트 존재 여부 확인<br>2. 핵심 컬럼 존재 여부 점검<br>3. PANEL/ALERT 정합성 및 lag 구조 검증",
        "분석 결과 자체보다, 결과를 뒷받침하는 데이터 구조와 계산 일관성을 확인하는 검증 단계로 해석한다."
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
