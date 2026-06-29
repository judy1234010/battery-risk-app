import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="망보는사람들 공급망 리스크 대시보드", page_icon="📊", layout="wide")

CHAIN_ORDER = ["납산배터리군", "리튬이온배터리군"]
CHAIN_COLOR_MAP = {"납산배터리군": "#1f77b4", "리튬이온배터리군": "#d62728"}
GRADE_COLOR_MAP = {"낮음": "#4C78A8", "보통": "#F2C14E", "높음": "#F28E2B", "매우높음": "#E15759", "해석유보": "#9D9D9D"}
AXIS_COLOR_MAP = {"가격리스크점수": "#4C78A8", "수급리스크점수": "#E15759", "물류리스크점수": "#59A14F", "정책이벤트리스크점수": "#B07AA1"}

SHEET_DESC = {
    "COUNTRY_MONTHLY": "체인·월별 국가 단위 수입 비중, 기본위험, 보정점수, 최종보정점수를 담은 시트입니다. 국가/공급선 상세분석과 대체국 추천의 기반입니다.",
    "PANEL_MONTHLY": "체인·월별 핵심 집계 시트입니다. 가격·수급·물류·정책이벤트 리스크, 최종위험점수(원점수), 상대위험지수(정규화값)가 들어 있습니다.",
    "ALERT_RESULT": "체인·월별 경보 결과 시트입니다. 최종경보등급, 보정사유, 대체조달가능성, 상대적 우선관리대상 여부가 들어 있습니다.",
    "체인별 비교표": "체인별 평균 수준, 분위수(Q25/Q50/Q75), 우선관리 비중, 30점 이상 개월수 등을 요약한 비교 시트입니다.",
    "ENTROPY_WEIGHT": "체인별 엔트로피 가중치(EWM) 시트입니다. 가격, 수급기초, 수급최종, 최종결합, 대체조달 가중치가 포함됩니다.",
    "NOMALIZATION_CHECK": "정규화와 분위수 계산 검증용 요약 시트입니다. min, q25, median, q75, max가 정리되어 있습니다.",
    "NOMALIZATION_AUDIT": "정규화 산출을 더 자세히 검증하는 감사용 시트입니다.",
    "METHOD_GUIDE": "본 모델의 데이터 처리 및 계산 방법을 서술한 방법론 시트입니다.",
    "SIGNAL_BASE": "선행신호 분석의 원천 변수 시트입니다. HHI, 상위1국의존도, 수입국수, 환율정규화 등 후보변수가 포함됩니다.",
    "SIGNAL_LAG_TABLE": "선행신호 후보변수의 lag 1~6개월 시차 테이블입니다.",
    "LEAD_SIGNAL_LAG_DETAIL": "체인·지표·lag별 상관계수 상세 시트입니다.",
    "LEAD_SIGNAL_LAG_COMPARE": "지표별 최적 lag와 절대상관을 비교한 요약 시트입니다.",
}

def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def safe_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def fmt_num(v, digits=2):
    if pd.isna(v):
        return "-"
    return f"{float(v):,.{digits}f}"

def fmt_pct(v, digits=2):
    if pd.isna(v):
        return "-"
    return f"{float(v):,.{digits}f}%"

def risk_label(score):
    if pd.isna(score):
        return "해석유보"
    if score >= 75:
        return "매우높음"
    if score >= 50:
        return "높음"
    if score >= 25:
        return "보통"
    return "낮음"

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
                    y = int(parts[0]); m = int(parts[1])
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
            y = int(s[:4]); m = int(s[4:])
            if 1 <= m <= 12:
                return f"{y:04d}-{m:02d}"
        return s
    return None

def ensure_ym_column(df):
    df = df.copy()
    cols = df.columns.tolist()
    if "연월_키" in cols:
        dt = pd.to_datetime(df["연월_키"], errors="coerce")
        df["연월"] = dt.dt.strftime("%Y-%m")
        return df
    if "연도" in cols and "월" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r["연도"], r["월"]), axis=1)
        return df
    if "Year" in cols and "Month" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r["Year"], r["Month"]), axis=1)
        return df
    if "연" in cols and "월" in cols:
        df["연월"] = df.apply(lambda r: safe_ym(None, r["연"], r["월"]), axis=1)
        return df
    if "연월" in cols:
        df["연월"] = df["연월"].apply(safe_ym)
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

def style_metric_container():
    st.markdown("""
    <style>
    div[data-testid="stMetric"] {
        background-color: rgba(255,255,255,0.04);
        border: 1px solid rgba(128,128,128,0.25);
        padding: 12px 14px;
        border-radius: 12px;
    }
    .small-note {
        font-size: 0.92rem;
        line-height: 1.6;
    }
    </style>
    """, unsafe_allow_html=True)

def info_box(title, bullets):
    if isinstance(bullets, str):
        bullets = [bullets]
    items = "".join([f"<li>{b}</li>" for b in bullets])
    st.markdown(
        f"""
        <div style="border:1px solid rgba(128,128,128,0.25); border-radius:12px; padding:14px 16px; margin-bottom:14px;">
            <div style="font-weight:700; margin-bottom:8px;">{title}</div>
            <ul class="small-note" style="margin-top:0; padding-left:18px;">{items}</ul>
        </div>
        """,
        unsafe_allow_html=True
    )

def build_download_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=str(sname)[:31], index=False)
    return output.getvalue()

def month_chain_slice(df, month=None, chain=None):
    tmp = df.copy()
    if month and "연월" in tmp.columns:
        tmp = tmp[tmp["연월"] == month].copy()
    if chain and "체인구분" in tmp.columns:
        tmp = tmp[tmp["체인구분"] == chain].copy()
    return tmp

def longest_run(mask_series):
    cur = 0
    best = 0
    for x in mask_series.fillna(False).astype(bool).tolist():
        cur = cur + 1 if x else 0
        best = max(best, cur)
    return best

def add_priority_metrics(alert_df, panel_df, compare_df):
    df = alert_df.copy()
    if "최종위험점수(원점수)" not in df.columns and "최종위험점수(원점수)" in panel_df.columns:
        df = df.merge(panel_df[["연월", "체인구분", "최종위험점수(원점수)"]], on=["연월", "체인구분"], how="left")
    keep = [c for c in ["체인구분", "상대위험지수_Q25", "상대위험지수_Q50", "상대위험지수_Q75",
                         "최종위험점수(원점수)_Q25", "최종위험점수(원점수)_Q50", "최종위험점수(원점수)_Q75"] if c in compare_df.columns]
    if keep:
        df = df.merge(compare_df[keep], on="체인구분", how="left")
    alt_q = df.groupby("체인구분")["대체조달가능성_점수"].quantile(0.25).rename("대체조달가능성_Q25").reset_index()
    df = df.merge(alt_q, on="체인구분", how="left")
    df["Q75_초과폭"] = safe_numeric(df["상대위험지수(정규화값)"]) - safe_numeric(df["상대위험지수_Q75"])
    df["Q25_미달폭"] = safe_numeric(df["대체조달가능성_Q25"]) - safe_numeric(df["대체조달가능성_점수"])
    df["우선관리강도"] = df[["Q75_초과폭", "Q25_미달폭"]].clip(lower=0).sum(axis=1)
    return df

def aggregate_country_view(country_df):
    if country_df.empty:
        return country_df
    tmp = country_df.copy()
    weight = safe_numeric(tmp["국가수입금액"]) if "국가수입금액" in tmp.columns else safe_numeric(tmp["국가별수입비중"]).fillna(0)
    tmp["_w"] = weight.fillna(0)
    group_cols = [c for c in ["국가코드", "국가명", "지역권"] if c in tmp.columns]
    agg_rows = []
    for keys, g in tmp.groupby(group_cols, dropna=False):
        row = {}
        if not isinstance(keys, tuple):
            keys = (keys,)
        for c, v in zip(group_cols, keys):
            row[c] = v
        w = g["_w"].sum()
        row["FTA여부"] = "Y" if ("FTA여부" in g.columns and (g["FTA여부"] == "Y").any()) else "N"
        row["상위공급국여부"] = "Y" if ("상위공급국여부" in g.columns and (g["상위공급국여부"] == "Y").any()) else "N"
        for c in ["국가수입중량", "국가수입금액", "중량비중", "금액비중", "국가별수입비중", "지역권별수입비중"]:
            if c in g.columns:
                row[c] = safe_numeric(g[c]).sum()
        for c in ["기본평가점수", "최종보정점수", "집중도기여도"]:
            if c in g.columns:
                val = np.average(safe_numeric(g[c]).fillna(0), weights=np.maximum(g["_w"], 1e-9)) if len(g) else np.nan
                row[c] = val
        if "위험점수출처" in g.columns:
            row["위험점수출처"] = ", ".join(sorted(g["위험점수출처"].dropna().astype(str).unique().tolist())[:2])
        if "비고" in g.columns:
            row["비고"] = ", ".join(sorted(g["비고"].dropna().astype(str).unique().tolist())[:2])
        row["국가공급선_최종판정"] = risk_label(row.get("최종보정점수", np.nan))
        row["국가기본위험_평가등급"] = risk_label(row.get("기본평가점수", np.nan))
        agg_rows.append(row)
    out = pd.DataFrame(agg_rows)
    if "국가별수입비중" in out.columns:
        out = out.sort_values("국가별수입비중", ascending=False).reset_index(drop=True)
    return out

def action_message(axis_reason, risk_value, alt_value, priority):
    msgs = []
    if axis_reason == "수급":
        msgs.append("공급국 집중도가 높게 작동한 달일 가능성이 커서 상위 공급국 의존도와 HHI를 우선 점검합니다.")
    elif axis_reason == "가격":
        msgs.append("원재료 또는 환율 변동이 최종위험점수에 크게 반영된 달입니다. 가격헤지·판가연동 조항 검토가 우선입니다.")
    elif axis_reason == "물류":
        msgs.append("글로벌 공급망 병목 또는 운송 차질의 영향이 큰 구간입니다. 재고일수와 운송경로 다변화 검토가 우선입니다.")
    elif axis_reason == "정책이벤트":
        msgs.append("정책·지정학 이벤트 충격이 반영된 구간입니다. 규제 공지, 제재, 통상정책 모니터링을 강화해야 합니다.")
    if alt_value <= 25:
        msgs.append("대체조달가능성이 낮아 단기 대체선 발굴, 계약 분산, FTA 활용 가능국 우선 검토가 필요합니다.")
    if priority == "Y":
        msgs.append("상대위험지수가 체인 상위구간(Q75 이상)이고 대체조달가능성이 하위구간(Q25 이하)이므로 우선관리 대상으로 해석합니다.")
    else:
        msgs.append("즉시 우선관리 대상은 아니지만, 주도 축과 조달여건을 함께 보며 추세를 모니터링하는 것이 적절합니다.")
    return " ".join(msgs)

@st.cache_data(show_spinner=False)
def load_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    raw = {s: clean_columns(pd.read_excel(uploaded_file, sheet_name=s)) for s in xls.sheet_names}

    def g(name):
        return raw.get(name, pd.DataFrame())

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

    for df in [country, panel, alert, compare, entropy, signal_base, signal_lag, lead_detail, lead_compare]:
        if "체인구분" in df.columns:
            df["체인구분"] = df["체인구분"].astype(str).str.strip()

    num_candidates = set()
    for df in [country, panel, alert, compare, entropy, signal_base, signal_lag, lead_detail, lead_compare, norm_check, norm_audit]:
        for c in df.columns:
            if any(tok in str(c) for tok in ["점수", "비중", "의존도", "HHI", "금액", "물량", "가격", "GSCPI", "TPU", "lag", "q25", "q75", "median", "max", "min", "_corr", "관측치", "개월수"]):
                num_candidates.add(c)
    for df in [country, panel, alert, compare, entropy, signal_base, signal_lag, lead_detail, lead_compare, norm_check, norm_audit]:
        for c in [x for x in num_candidates if x in df.columns]:
            df[c] = safe_numeric(df[c])

    if "보정사유" in alert.columns:
        alert["보정사유"] = alert["보정사유"].replace({"정책": "정책이벤트"})
    if "비고" in alert.columns:
        alert["비고"] = alert["비고"].astype(str).str.replace("정책 축", "정책이벤트 축", regex=False)
    alert = add_priority_metrics(alert, panel, compare)

    return {
        "sheets": raw, "country": country, "panel": panel, "alert": alert, "compare": compare,
        "entropy": entropy, "norm_check": norm_check, "norm_audit": norm_audit, "method": method,
        "signal_base": signal_base, "signal_lag": signal_lag, "lead_detail": lead_detail, "lead_compare": lead_compare
    }

style_metric_container()
uploaded_file = st.sidebar.file_uploader("최종 확정본 엑셀 업로드", type=["xlsx"])
if uploaded_file is None:
    st.title("망보는사람들 공급망 리스크 대시보드")
    st.info("최종 확정본 엑셀 파일을 업로드하면 메뉴가 활성화됩니다.")
    st.stop()

data = load_excel(uploaded_file)
panel = data["panel"]; alert = data["alert"]; country = data["country"]; compare = data["compare"]
entropy = data["entropy"]; norm_check = data["norm_check"]; norm_audit = data["norm_audit"]; method = data["method"]
lead_detail = data["lead_detail"]; lead_compare = data["lead_compare"]; signal_base = data["signal_base"]; signal_lag = data["signal_lag"]

month_list = get_month_list(panel if not panel.empty else alert)
chain_list = get_chain_list(panel if not panel.empty else alert)
if not month_list or not chain_list:
    st.error("연월 또는 체인구분을 읽지 못했습니다. 최종 엑셀 구조를 확인해주세요.")
    st.stop()

menu = st.sidebar.radio("메뉴 선택", [
    "1. 종합 상황판",
    "2. 체인별 심층 분석",
    "3. 국가/공급선 상세 분석",
    "4. 충격 원인 추적",
    "5. 선행 신호 후보 탐지",
    "6. 기업 대응 우선순위 추천 / 시뮬레이터",
    "7. 대체국 추천 시스템",
    "8. 원천데이터 탐색 / 다운로드",
    "9. 데이터 검증 / 방법론",
])

needs_month = menu in ["1. 종합 상황판", "3. 국가/공급선 상세 분석", "4. 충격 원인 추적", "7. 대체국 추천 시스템", "9. 데이터 검증 / 방법론"]
needs_chain = menu in ["2. 체인별 심층 분석", "3. 국가/공급선 상세 분석", "4. 충격 원인 추적", "5. 선행 신호 후보 탐지", "7. 대체국 추천 시스템", "9. 데이터 검증 / 방법론"]
selected_month = st.sidebar.selectbox("기준 연월", month_list, index=len(month_list)-1) if needs_month else None
selected_chain = st.sidebar.selectbox("기준 체인", chain_list, index=0) if needs_chain else None
st.sidebar.caption("연월은 모두 YYYY-MM 형식으로 표시됩니다.")

panel_month = month_chain_slice(panel, selected_month)
alert_month = month_chain_slice(alert, selected_month)
panel_chain = month_chain_slice(panel, None, selected_chain).sort_values("연월") if selected_chain else pd.DataFrame()
alert_chain = month_chain_slice(alert, None, selected_chain).sort_values("연월") if selected_chain else pd.DataFrame()
compare_chain = compare[compare["체인구분"] == selected_chain].copy() if selected_chain else pd.DataFrame()

if menu == "1. 종합 상황판":
    st.header("1. 종합 상황판")
    info_box("핵심 해석 원칙", [
        "최종위험점수(원점수)는 가격·수급·물류·정책이벤트 4개 축을 체인별 엔트로피 가중치(EWM)로 결합한 결과입니다. 즉, 단순 평균이 아니라 해당 체인에서 변동성과 구분력이 큰 축이 더 크게 반영됩니다.",
        "상대위험지수는 최종위험점수(원점수)를 같은 체인 내부의 과거 분포 기준으로 0~100 범위로 정규화한 상대지표입니다. 0은 무위험이 아니라 체인 내부 상대적 저점, 100은 절대 최대위험이 아니라 체인 내부 상대적 고점을 의미합니다.",
        "우선관리대상 Y는 상대위험지수가 체인별 상위 사분위(Q75) 이상이면서 동시에 대체조달가능성 점수가 체인별 하위 사분위(Q25) 이하인 경우입니다. 따라서 위험 수준이 높고, 동시에 대체조달 여건이 취약한 경우만 Y로 분류됩니다.",
        "Q75 초과폭은 현재 상대위험지수가 체인 상위 경계선(Q75)을 얼마나 넘어섰는지, Q25 미달폭은 대체조달가능성 점수가 체인 하위 경계선(Q25)보다 얼마나 더 낮은지를 뜻합니다. 두 값이 클수록 같은 Y 안에서도 관리 강도가 더 크다고 해석합니다."
    ])
    rows = []
    for chain in chain_list:
        p = panel_month[panel_month["체인구분"] == chain]
        a = alert_month[alert_month["체인구분"] == chain]
        c = compare[compare["체인구분"] == chain]
        if p.empty or a.empty or c.empty:
            continue
        p = p.iloc[0]; a = a.iloc[0]; c = c.iloc[0]
        rows.append({
            "체인구분": chain,
            "최종위험점수(원점수)": p["최종위험점수(원점수)"],
            "상대위험지수(정규화값)": a["상대위험지수(정규화값)"],
            "최종경보등급": a["최종경보등급"],
            "상대적_우선관리대상": a["상대적_우선관리대상"],
            "대체조달가능성_점수": a["대체조달가능성_점수"],
            "Q75_초과폭": a["Q75_초과폭"],
            "Q25_미달폭": a["Q25_미달폭"],
            "우선관리강도": a["우선관리강도"],
            "보정사유": a["보정사유"],
            "상대위험지수_Q75": c["상대위험지수_Q75"],
            "대체조달가능성_Q25": a["대체조달가능성_Q25"]
        })
    summary = pd.DataFrame(rows)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader(f"기준 연월: {selected_month}")
        st.caption("체인별로 현재 수준, 상대적 위치, 우선관리 여부를 한 번에 확인합니다.")
    with c2:
        st.markdown("**산식 요약**  \n최종위험점수(원점수) = 가격리스크×가중치 + 수급리스크×가중치 + 물류리스크×가중치 + 정책이벤트리스크×가중치")

    for chain in chain_list:
        row = summary[summary["체인구분"] == chain]
        if row.empty:
            continue
        r = row.iloc[0]
        st.markdown(f"#### {chain}")
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("최종위험점수(원점수)", fmt_num(r["최종위험점수(원점수)"]))
        m2.metric("상대위험지수", fmt_num(r["상대위험지수(정규화값)"]), r["최종경보등급"])
        m3.metric("우선관리대상", r["상대적_우선관리대상"], "Y=고위험·저대체조달")
        m4.metric("관리강도", fmt_num(r["우선관리강도"]), "Q75 초과 + Q25 미달")
        m5.metric("Q75 초과폭", fmt_num(r["Q75_초과폭"]), f"기준 Q75={fmt_num(r['상대위험지수_Q75'])}")
        m6.metric("Q25 미달폭", fmt_num(r["Q25_미달폭"]), f"대체조달={fmt_num(r['대체조달가능성_점수'])}")

        row_panel = panel_month[panel_month["체인구분"] == chain].iloc[0]
        contrib = pd.DataFrame({
            "축": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"],
            "점수": [row_panel["가격리스크점수"], row_panel["수급리스크점수"], row_panel["물류리스크점수"], row_panel["정책이벤트리스크점수"]]
        })
        fig = px.bar(contrib, x="축", y="점수", color="축", color_discrete_map=AXIS_COLOR_MAP, title=f"{chain} · 최종위험점수 구성축")
        fig.update_layout(height=320, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 우선관리 사분면")
    fig2 = px.scatter(
        summary, x="대체조달가능성_점수", y="상대위험지수(정규화값)", color="체인구분",
        size=np.maximum(summary["최종위험점수(원점수)"], 1), text="상대적_우선관리대상",
        color_discrete_map=CHAIN_COLOR_MAP, hover_data=["Q75_초과폭", "Q25_미달폭", "우선관리강도", "보정사유"]
    )
    for _, r in summary.iterrows():
        fig2.add_hline(y=r["상대위험지수_Q75"], line_dash="dot", line_color=CHAIN_COLOR_MAP.get(r["체인구분"], "#999"))
        fig2.add_vline(x=r["대체조달가능성_Q25"], line_dash="dot", line_color=CHAIN_COLOR_MAP.get(r["체인구분"], "#999"))
    if not summary.empty:
        fig2.add_annotation(x=summary["대체조달가능성_점수"].min(), y=summary["상대위험지수(정규화값)"].max()+5, text="좌상단 = Y 후보 영역(고위험·저대체조달)", showarrow=False)
        fig2.add_annotation(x=summary["대체조달가능성_점수"].max(), y=summary["상대위험지수(정규화값)"].min()-5, text="우하단 = N 후보 영역(저위험·고대체조달)", showarrow=False)
    fig2.update_layout(height=420)
    st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 체인별 종합표")
    st.dataframe(summary[["체인구분","최종위험점수(원점수)","상대위험지수(정규화값)","최종경보등급","상대적_우선관리대상","보정사유","Q75_초과폭","Q25_미달폭","우선관리강도"]], use_container_width=True, hide_index=True)

elif menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")
    info_box("핵심 해석 원칙", [
        "이 메뉴는 특정 연월의 단면이 아니라 체인 전체 기간의 구조를 보는 메뉴입니다. 그래서 기준 연월 선택은 두지 않고, 체인별 평균·분위수·월별 추이를 중심으로 해석합니다.",
        "최종위험점수(원점수)와 상대위험지수의 추이가 비슷하게 보이는 것은 오류가 아니라 체인 내부에서 원점수를 0~100으로 선형 정규화했기 때문입니다. 즉, 방향과 고점·저점의 위치는 같고, 스케일만 다릅니다.",
        "체인 간 비교에서는 상대위험지수만 보면 안 됩니다. 같은 30점이라도 체인별 원점수 분포와 Q25·Q50·Q75 경계값이 다르므로, 반드시 원점수 분위수와 함께 읽어야 합니다."
    ])
    st.subheader("체인별 기준선 비교")
    st.dataframe(compare, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        qdf = compare[["체인구분","상대위험지수_Q25","상대위험지수_Q50","상대위험지수_Q75"]].melt(id_vars="체인구분", var_name="구간", value_name="값")
        fig = px.bar(qdf, x="구간", y="값", color="체인구분", barmode="group", color_discrete_map=CHAIN_COLOR_MAP, title="체인별 상대위험지수 분위수")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        qdf2 = compare[["체인구분","최종위험점수(원점수)_Q25","최종위험점수(원점수)_Q50","최종위험점수(원점수)_Q75"]].melt(id_vars="체인구분", var_name="구간", value_name="값")
        fig2 = px.bar(qdf2, x="구간", y="값", color="체인구분", barmode="group", color_discrete_map=CHAIN_COLOR_MAP, title="체인별 원점수 분위수")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("체인별 월별 추이")
    tab1, tab2 = st.tabs(["최종위험점수(원점수) 추이", "상대위험지수 추이"])
    with tab1:
        fig3 = px.line(panel, x="연월", y="최종위험점수(원점수)", color="체인구분", markers=True, color_discrete_map=CHAIN_COLOR_MAP,
                       title="체인별 최종위험점수(원점수) 추이")
        st.plotly_chart(fig3, use_container_width=True)
    with tab2:
        fig4 = px.line(panel, x="연월", y="상대위험지수(정규화값)", color="체인구분", markers=True, color_discrete_map=CHAIN_COLOR_MAP,
                       title="체인별 상대위험지수 추이")
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("4대 리스크 평균 비교")
    axis_cols = ["평균_가격리스크점수","평균_수급리스크점수","평균_물류리스크점수","평균_정책이벤트리스크점수"]
    map_name = {"평균_가격리스크점수":"가격리스크점수","평균_수급리스크점수":"수급리스크점수","평균_물류리스크점수":"물류리스크점수","평균_정책이벤트리스크점수":"정책이벤트리스크점수"}
    a = compare[["체인구분"] + axis_cols].melt(id_vars="체인구분", var_name="항목", value_name="값")
    a["항목"] = a["항목"].map(map_name)
    fig5 = px.bar(a, x="항목", y="값", color="체인구분", barmode="group", color_discrete_map=CHAIN_COLOR_MAP, title="체인별 4대 리스크 평균 비교")
    st.plotly_chart(fig5, use_container_width=True)

elif menu == "3. 국가/공급선 상세 분석":
    st.header("3. 국가/공급선 상세 분석")
    info_box("핵심 해석 원칙", [
        "이 메뉴는 선택한 체인과 연월의 국가별 공급선 구조를 보여줍니다. 동일 국가가 여러 HS코드로 분리되어 입력된 경우가 있어, 앱에서는 국가 단위로 다시 합산해 보여줍니다. 따라서 막대가 둘로 쪼개져 보이던 오류를 방지합니다.",
        "표의 '국가공급선 최종판정'은 특정 월·체인 안에서 그 국가 공급선의 기본위험과 집중도 보정, FTA 보정 등을 합친 뒤 계산된 최종보정점수의 수준을 의미합니다. 즉, 단순 국가위험이 아니라 공급선 관점의 종합 판정입니다.",
        "국가기본위험 평가등급은 보정 전 기본위험 수준, 국가공급선 최종판정은 비중과 조달구조 보정까지 반영한 최종 수준으로 이해하면 됩니다."
    ])
    ctry = month_chain_slice(country, selected_month, selected_chain)
    ctry_agg = aggregate_country_view(ctry)
    if ctry_agg.empty:
        st.warning("선택한 조건에 해당하는 국가 데이터가 없습니다.")
    else:
        top_n = st.slider("상위 국가 표시 수", 5, min(30, len(ctry_agg)), min(15, len(ctry_agg)))
        show = ctry_agg.head(top_n).copy()
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(show.sort_values("국가별수입비중"), x="국가별수입비중", y="국가명", orientation="h",
                         color="국가공급선_최종판정", color_discrete_map=GRADE_COLOR_MAP,
                         title=f"{selected_month} {selected_chain} 국가별 수입비중")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.scatter(show, x="최종보정점수", y="국가별수입비중",
                              color="FTA여부", size=np.maximum(show["국가수입금액"].fillna(0), 1),
                              hover_data=["국가명","지역권","국가기본위험_평가등급","국가공급선_최종판정"],
                              title="국가별 최종보정점수 vs 수입비중")
            fig2.update_xaxes(title="최종보정점수(낮을수록 상대적으로 안정)")
            fig2.update_yaxes(title="국가별 수입비중")
            st.plotly_chart(fig2, use_container_width=True)
        out_cols = [c for c in ["국가명","지역권","FTA여부","국가별수입비중","국가수입금액","기본평가점수","국가기본위험_평가등급","최종보정점수","국가공급선_최종판정","상위공급국여부","비고"] if c in show.columns]
        st.dataframe(show[out_cols], use_container_width=True, hide_index=True)

elif menu == "4. 충격 원인 추적":
    st.header("4. 충격 원인 추적")
    info_box("핵심 해석 원칙", [
        "충격 원인은 먼저 4대 리스크 축의 절대수준을 보고, 그 다음 최종위험점수(원점수), 마지막으로 상대위험지수의 체인 내 위치를 봅니다.",
        "보정사유는 해당 월에 상대적으로 두드러진 축을 보여주는 간단한 라벨입니다. 실제 해석은 가격·수급·물류·정책이벤트 4개 축 수치를 함께 보고 판단해야 합니다.",
        "같은 낮음/보통/높음 등급이라도 체인별 기준선이 다르므로, 아래의 Q25·Q50·Q75 숫자는 '이 체인에서 어디부터 상위구간인가'를 알려주는 경계값으로 읽으면 됩니다."
    ])
    prow = panel_month[panel_month["체인구분"] == selected_chain]
    arow = alert_month[alert_month["체인구분"] == selected_chain]
    crow = compare_chain
    if prow.empty or arow.empty or crow.empty:
        st.warning("선택한 조건의 데이터가 없습니다.")
    else:
        prow = prow.iloc[0]; arow = arow.iloc[0]; crow = crow.iloc[0]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("최종위험점수(원점수)", fmt_num(prow["최종위험점수(원점수)"]))
        c2.metric("상대위험지수", fmt_num(arow["상대위험지수(정규화값)"]), arow["최종경보등급"])
        c3.metric("보정사유", str(arow["보정사유"]))
        c4.metric("우선관리대상", str(arow["상대적_우선관리대상"]))
        comp = pd.DataFrame({"축":["가격리스크점수","수급리스크점수","물류리스크점수","정책이벤트리스크점수"],
                             "값":[prow["가격리스크점수"],prow["수급리스크점수"],prow["물류리스크점수"],prow["정책이벤트리스크점수"]]})
        fig = px.bar(comp, x="축", y="값", color="축", color_discrete_map=AXIS_COLOR_MAP, title=f"{selected_month} {selected_chain} 4대 리스크")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown(
            f"- 상대위험지수 분위수 경계: **Q25 ({fmt_num(crow['상대위험지수_Q25'])}) / Q50 ({fmt_num(crow['상대위험지수_Q50'])}) / Q75 ({fmt_num(crow['상대위험지수_Q75'])})**  \n"
            f"- 원점수 분위수 경계: **Q25 ({fmt_num(crow['최종위험점수(원점수)_Q25'])}) / Q50 ({fmt_num(crow['최종위험점수(원점수)_Q50'])}) / Q75 ({fmt_num(crow['최종위험점수(원점수)_Q75'])})**"
        )

elif menu == "5. 선행 신호 후보 탐지":
    st.header("5. 선행 신호 후보 탐지")
    info_box("핵심 해석 원칙", [
        "이 메뉴는 연월을 바꿔도 달라지지 않는 구조 분석 화면입니다. 따라서 기준 연월 선택은 두지 않고, 체인별로 어떤 지표가 몇 개월 선행해 최종위험점수와 연관되는지를 봅니다.",
        "최적 lag는 '지표가 지금 움직였을 때 몇 개월 뒤 최종위험점수와 가장 강하게 연결되는가'를 의미합니다. 예를 들어 lag 2가 가장 크면, 해당 지표는 약 2개월 뒤 위험 수준과 가장 관련이 높았다는 뜻입니다.",
        "상관은 인과를 뜻하지는 않지만, 모니터링 우선순위와 경보 리드타임 설정에는 유용합니다. 즉, 최적 lag가 큰 지표는 조기경보 지표 후보, lag가 짧고 상관이 큰 지표는 즉시대응 지표 후보로 볼 수 있습니다."
    ])
    chain_compare = lead_compare[lead_compare["체인구분"] == selected_chain].copy()
    chain_detail = lead_detail[lead_detail["체인구분"] == selected_chain].copy()
    topn = st.slider("상위 후보 수", 5, min(20, len(chain_compare)) if len(chain_compare) > 0 else 5, 10)
    rank_df = chain_compare.sort_values("최적lag절대상관", ascending=False).head(topn)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(rank_df.sort_values("최적lag절대상관"), x="최적lag절대상관", y="지표", orientation="h", color="최적lag개월",
                     title=f"{selected_chain} 최적 lag 상위 후보")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        options = rank_df["지표"].tolist()
        sel_signal = st.selectbox("세부 시차를 볼 지표 선택", options if options else [""])
        if sel_signal and sel_signal in chain_detail["지표"].astype(str).tolist():
            tmp = chain_detail[chain_detail["지표"] == sel_signal].sort_values("lag개월")
            fig2 = px.line(tmp, x="lag개월", y="상관계수", markers=True, title=f"{selected_chain} · {sel_signal} lag별 상관")
            st.plotly_chart(fig2, use_container_width=True)
    st.dataframe(rank_df, use_container_width=True, hide_index=True)

    st.subheader("활용 가이드")
    if not rank_df.empty:
        g = rank_df.iloc[0]
        st.markdown(
            f"""
            - **가장 우선적으로 볼 지표**: `{g['지표']}`  
            - **권장 리드타임**: 약 **{int(g['최적lag개월'])}개월 전**부터 모니터링 강화  
            - **실무 해석**: 이 지표는 최종위험점수와의 절대상관이 **{fmt_num(g['최적lag절대상관'], 3)}**로 가장 크게 나타났습니다.  
            - **권장 활용**:  
              1. 월간 조기경보 보고서에 해당 지표를 선행지표로 별도 표기  
              2. 최적 lag 개월수만큼 앞서 임계치 이탈 여부를 점검  
              3. 동일 체인의 보정사유(가격/수급/물류/정책이벤트)와 함께 교차 확인  
            """
        )
elif menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.header("6. 기업 대응 우선순위 추천 / 시뮬레이터")
    info_box("핵심 해석 원칙", [
        "이 메뉴는 개별 연월을 보는 화면이 아니라, 전체 기간 중 어떤 월이 실제로 먼저 대응되어야 하는지를 우선순위로 정렬하는 메뉴입니다. 따라서 기준 연월 선택은 두지 않습니다.",
        "우선순위 추천은 상대위험지수, 대체조달가능성, Q75 초과폭, Q25 미달폭, 보정사유를 함께 사용합니다. 같은 Y라도 위험 초과가 큰지, 조달 취약이 큰지에 따라 대응 순서를 다르게 봐야 합니다.",
        "시뮬레이터는 특정 대응이 성공했을 때 우선관리 상태가 얼마나 완화되는지 보는 기능입니다. 예를 들어 가격헤지, 공급선 다변화, 물류 안정화, 정책충격 완화 등의 가정을 두고 원점수 또는 대체조달가능성을 조정해 볼 수 있습니다."
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
    show_cols = ["연월","체인구분","최종위험점수(원점수)","상대위험지수(정규화값)","대체조달가능성_점수","Q75_초과폭","Q25_미달폭","우선관리강도","상대적_우선관리대상","보정사유","추천우선순위"]
    st.dataframe(view_df[show_cols], use_container_width=True, hide_index=True)

    st.subheader("대응 시뮬레이터")
    pick_idx = st.selectbox("시뮬레이션 대상 행 선택", view_df.index.tolist(), format_func=lambda i: f"{view_df.loc[i,'연월']} | {view_df.loc[i,'체인구분']} | {view_df.loc[i,'보정사유']}")
    base = view_df.loc[pick_idx]
    action_type = st.selectbox("가정할 대응유형", ["가격충격 완화", "수급집중 완화", "물류차질 완화", "정책이벤트 완화", "대체조달선 확보"])
    if action_type == "대체조달선 확보":
        improve_alt = st.slider("대체조달가능성 개선폭", 0.0, 40.0, 10.0, 1.0)
        new_alt = min(100.0, float(base["대체조달가능성_점수"]) + improve_alt)
        new_rel = float(base["상대위험지수(정규화값)"])
    else:
        reduce_rel = st.slider("상대위험지수 완화폭", 0.0, 40.0, 10.0, 1.0)
        new_rel = max(0.0, float(base["상대위험지수(정규화값)"]) - reduce_rel)
        new_alt = float(base["대체조달가능성_점수"])

    q75 = float(base["상대위험지수_Q75"]) if pd.notna(base["상대위험지수_Q75"]) else 75.0
    q25 = float(base["대체조달가능성_Q25"]) if pd.notna(base["대체조달가능성_Q25"]) else 25.0
    new_y = "Y" if (new_rel >= q75 and new_alt <= q25) else "N"

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("현재 상대위험지수", fmt_num(base["상대위험지수(정규화값)"]))
    a2.metric("현재 대체조달가능성", fmt_num(base["대체조달가능성_점수"]))
    a3.metric("시뮬레이션 후 상대위험지수", fmt_num(new_rel))
    a4.metric("시뮬레이션 후 우선관리여부", new_y)

    st.markdown(action_message(base["보정사유"], base["상대위험지수(정규화값)"], base["대체조달가능성_점수"], base["상대적_우선관리대상"]))

elif menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")
    info_box("핵심 해석 원칙", [
        "이 메뉴의 핵심은 '어떤 국가가 상대적으로 덜 위험하고, 동시에 실제로 조달 후보로 검토할 만한가'를 찾는 것입니다.",
        "표의 국가기본위험 평가등급은 보정 전 기본위험 수준, 국가공급선 최종판정은 집중도·FTA 등을 반영한 공급선 관점의 종합 판정입니다.",
        "그래프는 x축에 최종보정점수, y축에 국가별 수입비중을 둡니다. 왼쪽 아래에 있을수록 상대적으로 안정적이면서 현재 비중이 낮아 대체선 검토 후보로 보기 쉽고, 오른쪽 위에 있을수록 위험과 집중이 동시에 큰 구간으로 해석할 수 있습니다."
    ])
    ctry = aggregate_country_view(month_chain_slice(country, selected_month, selected_chain))
    if ctry.empty:
        st.warning("선택한 조건에 해당하는 국가 데이터가 없습니다.")
    else:
        min_share = st.slider("최소 수입비중", 0.0, 20.0, 0.5, 0.5)
        prefer_fta = st.checkbox("FTA 국가 우선", value=False)
        cand = ctry[ctry["국가별수입비중"] >= min_share].copy()
        if prefer_fta:
            cand = cand[cand["FTA여부"] == "Y"].copy()
        cand["추천점수"] = (100 - cand["최종보정점수"]) * 0.5 + (100 - cand["국가별수입비중"]) * 0.2 + np.where(cand["FTA여부"] == "Y", 15, 0) + np.where(cand["상위공급국여부"] == "Y", -5, 5)
        cand = cand.sort_values(["추천점수", "최종보정점수"], ascending=[False, True]).copy()
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(cand.head(15).sort_values("추천점수"), x="추천점수", y="국가명", orientation="h", color="FTA여부", title="대체국 추천 상위 후보")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.scatter(cand.head(30), x="최종보정점수", y="국가별수입비중", color="FTA여부",
                              size=np.maximum(cand.head(30)["국가수입금액"].fillna(0), 1),
                              hover_data=["국가명","국가기본위험_평가등급","국가공급선_최종판정","추천점수"],
                              title="대체국 후보 분포")
            st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(cand[["국가명","지역권","FTA여부","국가별수입비중","최종보정점수","국가기본위험_평가등급","국가공급선_최종판정","추천점수"]], use_container_width=True, hide_index=True)

elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")
    info_box("핵심 해석 원칙", [
        "이 메뉴는 앱에 표시되는 값이 어떤 시트에서 왔는지 직접 확인하는 화면입니다.",
        "시트를 선택하면 아래에 해당 시트의 역할을 설명하고, 실제 데이터를 그대로 보여줍니다.",
        "다운로드한 파일은 현재 앱이 읽고 있는 최종 확정본 구조를 그대로 반영합니다."
    ])
    sheet_names = list(data["sheets"].keys())
    selected_sheet = st.selectbox("시트 선택", sheet_names)
    st.caption(SHEET_DESC.get(selected_sheet, "이 시트에 대한 별도 설명은 아직 등록되지 않았습니다."))
    sheet_df = ensure_ym_column(clean_columns(data["sheets"][selected_sheet].copy()))
    st.dataframe(sheet_df, use_container_width=True, hide_index=True)
    st.download_button("현재 시트 CSV 다운로드", sheet_df.to_csv(index=False).encode("utf-8-sig"), f"{selected_sheet}.csv", "text/csv")
    st.download_button("전체 시트 Excel 다운로드", build_download_excel(data["sheets"]), "battery_dashboard_export.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "9. 데이터 검증 / 방법론":
    st.header("9. 데이터 검증 / 방법론")
    info_box("핵심 해석 원칙", [
        "이 메뉴는 '값이 왜 이렇게 나왔는가'를 확인하는 검증용 화면입니다. 실무자는 여기서 체인별 가중치, 정규화 기준, 원점수와 상대위험지수의 연결관계를 확인할 수 있습니다.",
        "상대위험지수는 단독 숫자가 아니라, 해당 체인의 원점수 분포 안에서 현재 값이 어디쯤 위치하는지를 보여주는 보조 해석 지표입니다.",
        "따라서 실제 보고와 의사결정에서는 원점수, 상대위험지수, Q25·Q50·Q75 경계값을 함께 보는 것이 바람직합니다."
    ])
    tab1, tab2, tab3, tab4 = st.tabs(["방법론", "체인별 가중치", "정규화 검증", "선택 조건 감사"])
    with tab1:
        st.dataframe(method, use_container_width=True, hide_index=True)
    with tab2:
        stage = st.selectbox("가중치 단계 선택", entropy["단계"].dropna().astype(str).unique().tolist())
        w = entropy[entropy["단계"] == stage].copy()
        st.dataframe(w, use_container_width=True, hide_index=True)
        st.markdown("**가중치 합 점검**")
        st.dataframe(w.groupby(["단계","체인구분"], as_index=False)["가중치"].sum(), use_container_width=True, hide_index=True)
    with tab3:
        st.dataframe(norm_check, use_container_width=True, hide_index=True)
        if not norm_audit.empty:
            st.markdown("**추가 정규화 감사 시트**")
            st.dataframe(norm_audit, use_container_width=True, hide_index=True)
    with tab4:
        prow = panel_month[panel_month["체인구분"] == selected_chain]
        arow = alert_month[alert_month["체인구분"] == selected_chain]
        crow = compare_chain
        if prow.empty or arow.empty or crow.empty:
            st.warning("선택한 조건의 감사 데이터가 없습니다.")
        else:
            prow = prow.iloc[0]; arow = arow.iloc[0]; crow = crow.iloc[0]
            chain_panel = panel[panel["체인구분"] == selected_chain].copy()
            raw_min = chain_panel["최종위험점수(원점수)"].min()
            raw_max = chain_panel["최종위험점수(원점수)"].max()
            raw_val = prow["최종위험점수(원점수)"]
            calc_rel = ((raw_val - raw_min) / (raw_max - raw_min) * 100) if (pd.notna(raw_val) and pd.notna(raw_min) and pd.notna(raw_max) and raw_max != raw_min) else np.nan
            audit_df = pd.DataFrame([
                ["연월", selected_month], ["체인구분", selected_chain], ["최종위험점수(원점수)", raw_val],
                ["체인내 원점수 최소값", raw_min], ["체인내 원점수 최대값", raw_max],
                ["원점수 기준 재계산 상대위험지수", calc_rel], ["시트상 상대위험지수", arow["상대위험지수(정규화값)"]],
                ["상대위험지수 Q25", crow["상대위험지수_Q25"]], ["상대위험지수 Q50", crow["상대위험지수_Q50"]], ["상대위험지수 Q75", crow["상대위험지수_Q75"]],
                ["대체조달가능성 Q25", arow["대체조달가능성_Q25"]], ["현재 대체조달가능성", arow["대체조달가능성_점수"]],
                ["우선관리대상", arow["상대적_우선관리대상"]]
            ], columns=["항목", "값"])
            st.dataframe(audit_df, use_container_width=True, hide_index=True)
