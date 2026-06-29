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

DEFAULT_MONTH = "2025-12"
CHAIN_ORDER = ["납산배터리군", "리튬이온배터리군"]
RISK_COLOR_MAP = {
    "낮음": "#4C78A8",
    "보통": "#F2C14E",
    "높음": "#F28E2B",
    "매우높음": "#E15759",
    "해석유보": "#9D9D9D",
}
CHAIN_COLOR_MAP = {
    "납산배터리군": "#1f77b4",
    "리튬이온배터리군": "#d62728",
}
AXIS_COLOR_MAP = {
    "가격리스크점수": "#4C78A8",
    "수급리스크점수": "#E15759",
    "물류리스크점수": "#59A14F",
    "정책이벤트리스크점수": "#B07AA1",
}

# =========================================================
# 공통 유틸
# =========================================================
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
        return str(x)

    if isinstance(x, (float, np.floating)):
        return None

    return str(x)

def ensure_ym_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    cols = df.columns.tolist()

    if "연월_키" in cols:
        df["연월"] = pd.to_datetime(df["연월_키"], errors="coerce").dt.strftime("%Y-%m")
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
        converted = []
        for val in df["연월"]:
            if isinstance(val, (pd.Timestamp, str, int, np.integer)):
                converted.append(safe_ym(val))
            else:
                converted.append(None)
        converted = pd.Series(converted, index=df.index)
        if converted.notna().sum() > 0:
            df["연월"] = converted.fillna(df["연월"].astype(str))
        else:
            df["연월"] = df["연월"].astype(str)
        return df

    return df

def info_box(title, text):
    st.markdown(
        f"""
        <div style="border:1px solid rgba(0,0,0,0.08); border-radius:12px; padding:14px 16px; background:#f8fafc; margin-bottom:14px;">
            <div style="font-weight:700; margin-bottom:6px;">{title}</div>
            <div style="font-size:0.95rem; line-height:1.55;">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def get_month_list(df):
    if df is None or "연월" not in df.columns:
        return []
    vals = df["연월"].dropna().astype(str).unique().tolist()
    vals = [v for v in vals if len(v) == 7 and v[4] == "-"]
    return sorted(vals)

def get_chain_list(df):
    if df is None or "체인구분" not in df.columns:
        return []
    vals = df["체인구분"].dropna().astype(str).unique().tolist()
    return [x for x in CHAIN_ORDER if x in vals] + [x for x in vals if x not in CHAIN_ORDER]

def get_default_index(options, default_value=DEFAULT_MONTH):
    if not options:
        return 0
    if default_value in options:
        return options.index(default_value)
    return len(options) - 1

def build_download_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            if isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=str(sname)[:31], index=False)
    return output.getvalue()

def make_kpi_card(title, value, sub=None):
    sub_html = f"<div style='font-size:0.82rem;color:#6b7280;margin-top:4px;'>{sub}</div>" if sub else ""
    st.markdown(
        f"""
        <div style="border:1px solid rgba(0,0,0,0.08); border-radius:12px; padding:14px; background:white;">
            <div style="font-size:0.85rem; color:#6b7280;">{title}</div>
            <div style="font-size:1.55rem; font-weight:700; margin-top:6px;">{value}</div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

def risk_grade_from_quantiles(value, q25, q50, q75):
    if pd.isna(value):
        return "해석유보"
    if value >= q75:
        return "매우높음"
    if value >= q50:
        return "높음"
    if value >= q25:
        return "보통"
    return "낮음"

def longest_run(mask_series):
    max_run = 0
    current = 0
    for val in mask_series.fillna(False).astype(bool):
        current = current + 1 if val else 0
        max_run = max(max_run, current)
    return max_run

def get_weight_map(entropy_df, stage, chain):
    if entropy_df is None or entropy_df.empty:
        return {}
    tmp = entropy_df[(entropy_df["단계"] == stage) & (entropy_df["체인구분"] == chain)].copy()
    if tmp.empty:
        return {}
    return dict(zip(tmp["변수명"], safe_numeric(tmp["가중치"])))

def add_priority_metrics(alert_df, panel_df, compare_df):
    df = alert_df.copy()
    panel_cols = ["연월", "체인구분", "최종위험점수(원점수)"]
    if panel_df is not None and not panel_df.empty:
        df = df.merge(panel_df[panel_cols], on=["연월", "체인구분"], how="left")

    if compare_df is not None and not compare_df.empty:
        keep = [
            "체인구분",
            "상대위험지수_Q25",
            "상대위험지수_Q50",
            "상대위험지수_Q75",
            "최종위험점수(원점수)_Q25",
            "최종위험점수(원점수)_Q50",
            "최종위험점수(원점수)_Q75",
        ]
        keep = [c for c in keep if c in compare_df.columns]
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
    df["우선관리강도"] = df[["Q75_초과폭", "Q25_미달폭"]].clip(lower=0).sum(axis=1)
    return df

def get_chain_summary_text(compare_row):
    if compare_row is None or compare_row.empty:
        return ""
    row = compare_row.iloc[0]
    return (
        f"{row['체인구분']}은 평균 상대위험지수 {fmt_num(row['평균_상대위험지수'])}, "
        f"Q75 {fmt_num(row['상대위험지수_Q75'])}, "
        f"원점수 Q75 {fmt_num(row['최종위험점수(원점수)_Q75'])} 수준이다. "
        f"{row['종합_시사점']}"
    )

def month_chain_slice(df, month, chain=None):
    if df is None or df.empty:
        return df
    tmp = df[df["연월"] == month].copy() if month else df.copy()
    if chain:
        tmp = tmp[tmp["체인구분"] == chain].copy()
    return tmp

@st.cache_data(show_spinner=False)
def load_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    raw = {}
    for s in xls.sheet_names:
        raw[s] = clean_columns(pd.read_excel(uploaded_file, sheet_name=s))

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
    hs_summary = ensure_ym_column(g("HS_MONTHLY_SUMMARY"))
    market = ensure_ym_column(g("MARKET_INDEX"))
    gscpi = ensure_ym_column(g("GSCPI_INDEX"))
    tpu = ensure_ym_column(g("TPU_INDEX"))

    for df in [panel, alert, compare, entropy, signal_base, signal_lag, lead_detail, lead_compare, hs_summary, country]:
        if "체인구분" in df.columns:
            df["체인구분"] = df["체인구분"].astype(str).str.strip()

    num_cols = [
        "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
        "상대위험지수(정규화값)", "최종위험점수(원점수)", "대체조달가능성_점수",
        "상대위험지수_Q25", "상대위험지수_Q50", "상대위험지수_Q75",
        "최종위험점수(원점수)_Q25", "최종위험점수(원점수)_Q50", "최종위험점수(원점수)_Q75",
        "총수입금액", "총수입물량", "평균수입단가", "상위1국의존도", "상위3국집중도", "HHI", "수입국수",
        "경보점수기초", "국가보정합계", "환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화",
        "GSCPI_Norm", "TPU_INDEX_NORM", "fta_ratio", "지역권수"
    ]
    for df in [panel, alert, compare, country, signal_base, signal_lag, lead_detail, lead_compare, hs_summary, market, gscpi, tpu]:
        for c in [c for c in num_cols if c in df.columns]:
            df[c] = safe_numeric(df[c])

    alert = add_priority_metrics(alert, panel, compare)

    if "보정사유" in alert.columns:
        alert["보정사유"] = alert["보정사유"].replace({"정책": "정책이벤트"})
    if "비고" in alert.columns:
        alert["비고"] = alert["비고"].astype(str).str.replace("정책 축", "정책이벤트 축", regex=False)

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

# =========================================================
# 입력
# =========================================================
uploaded_file = st.sidebar.file_uploader("최종 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file is None:
    st.title("망보는사람들 공급망 리스크 대시보드")
    st.info("최종 확정본 엑셀 파일을 업로드하면 9개 메뉴가 활성화됩니다.")
    st.stop()

data = load_excel(uploaded_file)
panel = data["panel"]
alert = data["alert"]
country = data["country"]
compare = data["compare"]
entropy = data["entropy"]
norm_check = data["norm_check"]
method = data["method"]
lead_detail = data["lead_detail"]
lead_compare = data["lead_compare"]
signal_base = data["signal_base"]
signal_lag = data["signal_lag"]
hs_summary = data["hs_summary"]

month_list = get_month_list(panel if not panel.empty else alert)
chain_list = get_chain_list(panel if not panel.empty else alert)

if not month_list or not chain_list:
    st.error("연월 또는 체인구분 정보를 읽지 못했습니다. 최종 시트 구조를 확인해주세요.")
    st.stop()

selected_month = st.sidebar.selectbox("기준 연월", month_list, index=get_default_index(month_list))
selected_chain = st.sidebar.selectbox("기준 체인", chain_list, index=0)

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

st.sidebar.markdown("---")
st.sidebar.caption("연월 표기는 앱 전체에서 YYYY-MM 형식으로 통일됩니다.")
st.sidebar.caption("상대위험지수는 체인별 원점수를 0~100으로 재정규화한 상대지표입니다.")

panel_month = month_chain_slice(panel, selected_month)
alert_month = month_chain_slice(alert, selected_month)
panel_chain = panel[panel["체인구분"] == selected_chain].sort_values("연월").copy()
alert_chain = alert[alert["체인구분"] == selected_chain].sort_values("연월").copy()
compare_chain = compare[compare["체인구분"] == selected_chain].copy()

# =========================================================
# 1. 종합 상황판
# =========================================================
if menu == "1. 종합 상황판":
    st.header("1. 종합 상황판")
    info_box(
        "핵심 해석 원칙",
        "이 화면에서는 <b>최종위험점수(원점수)</b>와 <b>상대위험지수(정규화값)</b>를 반드시 함께 보여준다. "
        "같은 우선관리대상 Y라도 <b>Q75 초과폭</b>과 <b>Q25 미달폭</b>이 다를 수 있으므로, "
        "Y/N만이 아니라 관리 강도 차이까지 함께 해석해야 한다."
    )

    cards = []
    for chain in chain_list:
        arow = alert_month[alert_month["체인구분"] == chain]
        prow = panel_month[panel_month["체인구분"] == chain]
        if arow.empty or prow.empty:
            continue
        cards.append({
            "체인구분": chain,
            "최종위험점수(원점수)": prow["최종위험점수(원점수)"].iloc[0],
            "상대위험지수(정규화값)": arow["상대위험지수(정규화값)"].iloc[0],
            "최종경보등급": arow["최종경보등급"].iloc[0],
            "상대적_우선관리대상": arow["상대적_우선관리대상"].iloc[0],
            "Q75_초과폭": arow["Q75_초과폭"].iloc[0],
            "Q25_미달폭": arow["Q25_미달폭"].iloc[0],
            "대체조달가능성_점수": arow["대체조달가능성_점수"].iloc[0],
        })
    summary_df = pd.DataFrame(cards)

    cols = st.columns(len(chain_list))
    for i, chain in enumerate(chain_list):
        row = summary_df[summary_df["체인구분"] == chain]
        with cols[i]:
            if row.empty:
                st.warning(f"{chain} 데이터 없음")
            else:
                r = row.iloc[0]
                make_kpi_card(f"{chain} · 최종위험점수(원점수)", fmt_num(r["최종위험점수(원점수)"]))
                make_kpi_card(f"{chain} · 상대위험지수", fmt_num(r["상대위험지수(정규화값)"]), f"등급: {r['최종경보등급']}")
                make_kpi_card(
                    f"{chain} · 우선관리대상",
                    r["상대적_우선관리대상"],
                    f"Q75 초과폭 {fmt_num(r['Q75_초과폭'])} / Q25 미달폭 {fmt_num(r['Q25_미달폭'])}"
                )

    st.markdown("#### 체인별 핵심 비교")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(
            summary_df,
            x="체인구분",
            y=["최종위험점수(원점수)", "상대위험지수(정규화값)"],
            barmode="group",
            title=f"{selected_month} 체인별 원점수 vs 상대위험지수"
        )
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig2 = px.scatter(
            summary_df,
            x="대체조달가능성_점수",
            y="상대위험지수(정규화값)",
            color="체인구분",
            text="상대적_우선관리대상",
            size=np.maximum(summary_df["최종위험점수(원점수)"], 1),
            title=f"{selected_month} 우선관리 사분면",
            color_discrete_map=CHAIN_COLOR_MAP,
        )
        for chain in chain_list:
            qrow = compare[compare["체인구분"] == chain]
            if not qrow.empty:
                fig2.add_hline(y=qrow["상대위험지수_Q75"].iloc[0], line_dash="dot", line_color=CHAIN_COLOR_MAP.get(chain, "#999"))
        alt_q = alert.groupby("체인구분")["대체조달가능성_점수"].quantile(0.25).to_dict()
        for chain in chain_list:
            if chain in alt_q:
                fig2.add_vline(x=alt_q[chain], line_dash="dot", line_color=CHAIN_COLOR_MAP.get(chain, "#999"))
        st.plotly_chart(fig2, use_container_width=True)

    display_cols = [
        "체인구분", "최종위험점수(원점수)", "상대위험지수(정규화값)", "최종경보등급",
        "상대적_우선관리대상", "Q75_초과폭", "Q25_미달폭", "대체조달가능성_점수"
    ]
    st.dataframe(summary_df[display_cols], use_container_width=True, hide_index=True)

# =========================================================
# 2. 체인별 심층 분석
# =========================================================
elif menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")
    info_box(
        "핵심 해석 원칙",
        "체인 간 비교에서는 <b>상대위험지수만 단독 비교하면 안 된다.</b> "
        "체인별로 원점수 분포와 Q25·Q50·Q75 경계값이 다르므로, 반드시 "
        "<b>체인별 원점수 분위수와 상대위험지수 분위수</b>를 함께 본다."
    )

    st.markdown("#### 체인별 기준선 비교표")
    st.dataframe(compare, use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        qcols = ["체인구분", "상대위험지수_Q25", "상대위험지수_Q50", "상대위험지수_Q75"]
        qdf = compare[qcols].melt(id_vars="체인구분", var_name="분위", value_name="값")
        fig = px.bar(qdf, x="분위", y="값", color="체인구분", barmode="group", title="체인별 상대위험지수 분위수")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        qcols_raw = ["체인구분", "최종위험점수(원점수)_Q25", "최종위험점수(원점수)_Q50", "최종위험점수(원점수)_Q75"]
        qdf_raw = compare[qcols_raw].melt(id_vars="체인구분", var_name="분위", value_name="값")
        fig2 = px.bar(qdf_raw, x="분위", y="값", color="체인구분", barmode="group", title="체인별 원점수 분위수")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("#### 선택 체인 추이")
    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(x=panel_chain["연월"], y=panel_chain["최종위험점수(원점수)"], mode="lines+markers", name="최종위험점수(원점수)"))
    fig3.add_trace(go.Scatter(x=panel_chain["연월"], y=panel_chain["상대위험지수(정규화값)"], mode="lines+markers", name="상대위험지수", yaxis="y2"))
    fig3.update_layout(
        title=f"{selected_chain} 월별 추이",
        xaxis_title="연월",
        yaxis=dict(title="원점수"),
        yaxis2=dict(title="상대위험지수", overlaying="y", side="right"),
        legend=dict(orientation="h")
    )
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("#### 선택 체인 구성 리스크 평균")
    axis_cols = ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"]
    axis_mean = panel_chain[axis_cols].mean().reset_index()
    axis_mean.columns = ["축", "평균값"]
    fig4 = px.bar(axis_mean, x="축", y="평균값", color="축", color_discrete_map=AXIS_COLOR_MAP, title=f"{selected_chain} 4대 리스크 평균")
    st.plotly_chart(fig4, use_container_width=True)

    st.caption(get_chain_summary_text(compare_chain))

# =========================================================
# 3. 국가/공급선 상세 분석
# =========================================================
elif menu == "3. 국가/공급선 상세 분석":
    st.header("3. 국가/공급선 상세 분석")
    info_box(
        "핵심 해석 원칙",
        "국가 레벨에서는 <b>최종보정점수</b>, <b>국가별수입비중</b>, <b>FTA 여부</b>, <b>상위공급국 여부</b>를 함께 본다. "
        "체인별 수급리스크는 국가별 위험 분포와 집중도 구조가 결합되어 형성된다."
    )

    ctry = month_chain_slice(country, selected_month, selected_chain)
    if ctry.empty:
        st.warning("선택한 연월/체인에 해당하는 국가 데이터가 없습니다.")
    else:
        top_n = st.slider("상위 국가 표시 수", 5, min(30, len(ctry)), min(15, len(ctry)))
        rank_col = "국가별수입비중" if "국가별수입비중" in ctry.columns else "금액비중"
        show = ctry.sort_values(rank_col, ascending=False).head(top_n).copy()

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(
                show.sort_values(rank_col),
                x=rank_col, y="국가명", orientation="h",
                color="최종판정" if "최종판정" in show.columns else None,
                title=f"{selected_month} {selected_chain} 국가별 수입비중",
                color_discrete_map=RISK_COLOR_MAP
            )
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            if "최종보정점수" in show.columns:
                fig2 = px.scatter(
                    show, x=rank_col, y="최종보정점수",
                    color="FTA여부" if "FTA여부" in show.columns else None,
                    hover_data=["국가명", "지역권", "평가등급"] if "평가등급" in show.columns else ["국가명", "지역권"],
                    title="국가별 비중 vs 최종보정점수"
                )
                st.plotly_chart(fig2, use_container_width=True)

        detail_cols = [c for c in [
            "국가명", "지역권", "FTA여부", "국가별수입비중", "지역권별수입비중",
            "기본평가점수", "최종보정점수", "최종판정", "상위공급국여부", "비고"
        ] if c in show.columns]
        st.dataframe(show[detail_cols], use_container_width=True, hide_index=True)

# =========================================================
# 4. 충격 원인 추적
# =========================================================
elif menu == "4. 충격 원인 추적":
    st.header("4. 충격 원인 추적")
    info_box(
        "핵심 해석 원칙",
        "충격 원인 해석은 <b>4대 리스크 축 → 최종위험점수(원점수) → 상대위험지수</b> 순서로 읽는다. "
        "상대위험지수 0은 무위험이 아니라 해당 체인 내부에서 상대적으로 가장 낮은 위치일 뿐이다."
    )

    prow = panel_month[panel_month["체인구분"] == selected_chain]
    arow = alert_month[alert_month["체인구분"] == selected_chain]
    if prow.empty or arow.empty:
        st.warning("선택한 체인의 기준월 데이터가 없습니다.")
    else:
        prow = prow.iloc[0]
        arow = arow.iloc[0]

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            make_kpi_card("최종위험점수(원점수)", fmt_num(prow["최종위험점수(원점수)"]))
        with c2:
            make_kpi_card("상대위험지수", fmt_num(arow["상대위험지수(정규화값)"]), f"등급: {arow['최종경보등급']}")
        with c3:
            make_kpi_card("우선관리대상", arow["상대적_우선관리대상"], f"Q75 초과폭 {fmt_num(arow['Q75_초과폭'])}")
        with c4:
            make_kpi_card("대체조달가능성 점수", fmt_num(arow["대체조달가능성_점수"]), f"Q25 미달폭 {fmt_num(arow['Q25_미달폭'])}")

        comp_df = pd.DataFrame({
            "축": ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"],
            "값": [prow["가격리스크점수"], prow["수급리스크점수"], prow["물류리스크점수"], prow["정책이벤트리스크점수"]],
        })
        fig = px.bar(comp_df, x="축", y="값", color="축", color_discrete_map=AXIS_COLOR_MAP, title=f"{selected_month} {selected_chain} 4대 리스크 구성")
        st.plotly_chart(fig, use_container_width=True)

        stages = ["PANEL_PRICE", "PANEL_SUPPLY_BASE", "PANEL_SUPPLY_FINAL", "PANEL_FINAL", "ALERT_ALT_SOURCE"]
        tabs = st.tabs(["가격가중치", "수급기초가중치", "수급최종가중치", "최종가중치", "대체조달가중치"])
        for tab, stage in zip(tabs, stages):
            with tab:
                w = entropy[(entropy["단계"] == stage) & (entropy["체인구분"] == selected_chain)].copy()
                st.dataframe(w, use_container_width=True, hide_index=True)

        if not compare_chain.empty:
            q = compare_chain.iloc[0]
            st.markdown(
                f"""
                - 기준 체인 상대위험지수 분위수: Q25 {fmt_num(q['상대위험지수_Q25'])} / Q50 {fmt_num(q['상대위험지수_Q50'])} / Q75 {fmt_num(q['상대위험지수_Q75'])}  
                - 기준 체인 원점수 분위수: Q25 {fmt_num(q['최종위험점수(원점수)_Q25'])} / Q50 {fmt_num(q['최종위험점수(원점수)_Q50'])} / Q75 {fmt_num(q['최종위험점수(원점수)_Q75'])}
                """
            )

# =========================================================
# 5. 선행 신호 후보 탐지
# =========================================================
elif menu == "5. 선행 신호 후보 탐지":
    st.header("5. 선행 신호 후보 탐지")
    info_box(
        "핵심 해석 원칙",
        "선행신호 메뉴는 <b>체인별 원천 변수의 시차 상관</b>을 확인하는 화면이다. "
        "가공된 가격/수급/물류/정책 점수 자체가 아니라, HHI·상위1국의존도·수입국수·환율정규화·원재료 가격정규화 등 원천 후보변수의 lag 1~6개월 상관을 본다."
    )

    chain_detail = lead_detail[lead_detail["체인구분"] == selected_chain].copy()
    chain_compare = lead_compare[lead_compare["체인구분"] == selected_chain].copy()

    topn = st.slider("상위 후보 수", 5, min(20, len(chain_compare)) if len(chain_compare) > 0 else 5, 10)
    rank_df = chain_compare.sort_values("최적lag절대상관", ascending=False).head(topn)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(rank_df.sort_values("최적lag절대상관"),
                     x="최적lag절대상관", y="지표", orientation="h",
                     color="최적lag개월", title=f"{selected_chain} 최적 lag 상위 후보")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        sel_signal = st.selectbox("세부 시차를 볼 지표 선택", rank_df["지표"].tolist() if not rank_df.empty else [])
        if sel_signal:
            tmp = chain_detail[chain_detail["지표"] == sel_signal].sort_values("lag개월")
            fig2 = px.line(tmp, x="lag개월", y="상관계수", markers=True, title=f"{selected_chain} · {sel_signal} lag별 상관")
            st.plotly_chart(fig2, use_container_width=True)

    st.dataframe(rank_df, use_container_width=True, hide_index=True)
# =========================================================
# 6. 기업 대응 우선순위 추천 / 시뮬레이터
# =========================================================
elif menu == "6. 기업 대응 우선순위 추천 / 시뮬레이터":
    st.header("6. 기업 대응 우선순위 추천 / 시뮬레이터")
    info_box(
        "핵심 해석 원칙",
        "우선관리대상은 Y/N만으로 끝나지 않는다. "
        "같은 Y라도 <b>상대위험지수 Q75 초과폭</b>과 <b>대체조달가능성 Q25 미달폭</b>이 다르므로 "
        "실제 대응 우선순위는 강도 기준으로 재정렬해야 한다."
    )

    view_df = alert.copy()
    c1, c2, c3 = st.columns(3)
    with c1:
        chain_filter = st.multiselect("체인 선택", chain_list, default=chain_list)
    with c2:
        only_y = st.checkbox("우선관리대상 Y만 보기", value=False)
    with c3:
        sort_key = st.selectbox(
            "정렬 기준",
            ["우선관리강도", "상대위험지수(정규화값)", "최종위험점수(원점수)", "Q75_초과폭", "Q25_미달폭"]
        )

    view_df = view_df[view_df["체인구분"].isin(chain_filter)].copy()
    if only_y:
        view_df = view_df[view_df["상대적_우선관리대상"] == "Y"].copy()
    view_df = view_df.sort_values(sort_key, ascending=False)

    scenario_q75 = st.slider("가상 위험 기준(Q75 대체값)", 0.0, 100.0, 75.0, 0.5)
    scenario_q25 = st.slider("가상 조달 기준(Q25 대체값)", 0.0, 100.0, 25.0, 0.5)
    sim = view_df.copy()
    sim["시뮬레이션_우선관리"] = np.where(
        (safe_numeric(sim["상대위험지수(정규화값)"]) >= scenario_q75) &
        (safe_numeric(sim["대체조달가능성_점수"]) <= scenario_q25),
        "Y", "N"
    )

    fig = px.scatter(
        sim,
        x="대체조달가능성_점수",
        y="상대위험지수(정규화값)",
        color="체인구분",
        symbol="상대적_우선관리대상",
        size=np.maximum(safe_numeric(sim["최종위험점수(원점수)"]).fillna(0), 1),
        hover_data=["연월", "Q75_초과폭", "Q25_미달폭", "우선관리강도"],
        title="우선관리 사분면 및 강도 비교",
        color_discrete_map=CHAIN_COLOR_MAP,
    )
    fig.add_hline(y=scenario_q75, line_dash="dash", line_color="red")
    fig.add_vline(x=scenario_q25, line_dash="dash", line_color="blue")
    st.plotly_chart(fig, use_container_width=True)

    out_cols = [
        "연월", "체인구분", "최종위험점수(원점수)", "상대위험지수(정규화값)", "최종경보등급",
        "대체조달가능성_점수", "Q75_초과폭", "Q25_미달폭", "우선관리강도", "상대적_우선관리대상",
        "시뮬레이션_우선관리", "보정사유", "비고"
    ]
    st.dataframe(sim[out_cols], use_container_width=True, hide_index=True)

# =========================================================
# 7. 대체국 추천 시스템
# =========================================================
elif menu == "7. 대체국 추천 시스템":
    st.header("7. 대체국 추천 시스템")
    info_box(
        "핵심 해석 원칙",
        "대체국 추천은 국가별 <b>최종보정점수</b>가 낮고, <b>FTA 여부</b>가 유리하며, "
        "<b>현재 비중이 너무 과도하지 않은 국가</b>를 우선적으로 살핀다. "
        "이 화면은 실행 후보 탐색용이며, 계약·품질·운송조건은 별도 검토가 필요하다."
    )

    ctry = month_chain_slice(country, selected_month, selected_chain).copy()
    if ctry.empty:
        st.warning("선택한 연월/체인에 해당하는 국가 데이터가 없습니다.")
    else:
        min_share = st.slider("최소 수입비중(%)", 0.0, 20.0, 0.5, 0.5)
        prefer_fta = st.checkbox("FTA 국가 우선 보기", value=False)

        rank_col = "국가별수입비중" if "국가별수입비중" in ctry.columns else "금액비중"
        ctry = ctry[safe_numeric(ctry[rank_col]).fillna(0) >= min_share].copy()
        if prefer_fta and "FTA여부" in ctry.columns:
            ctry = ctry[ctry["FTA여부"] == "Y"].copy()

        ctry["추천점수"] = (
            (100 - safe_numeric(ctry["최종보정점수"]).fillna(100)) * 0.45 +
            (100 - safe_numeric(ctry[rank_col]).fillna(100)) * 0.20 +
            np.where(ctry.get("FTA여부", "N") == "Y", 20, 0) +
            np.where(ctry.get("상위공급국여부", "N") == "Y", -5, 5)
        )
        reco = ctry.sort_values(["추천점수", "최종보정점수"], ascending=[False, True]).copy()

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(
                reco.head(15).sort_values("추천점수"),
                x="추천점수", y="국가명", orientation="h",
                color="FTA여부" if "FTA여부" in reco.columns else None,
                title=f"{selected_month} {selected_chain} 대체국 추천 상위 15개"
            )
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.scatter(
                reco.head(30),
                x=rank_col, y="최종보정점수",
                color="FTA여부" if "FTA여부" in reco.columns else None,
                size=np.maximum(safe_numeric(reco.head(30)["추천점수"]).fillna(0), 1),
                hover_data=["국가명", "지역권", "평가등급"] if "평가등급" in reco.columns else ["국가명", "지역권"],
                title="대체국 후보 분포"
            )
            st.plotly_chart(fig2, use_container_width=True)

        out_cols = [c for c in [
            "국가명", "지역권", "FTA여부", rank_col, "기본평가점수", "최종보정점수",
            "최종판정", "상위공급국여부", "추천점수", "비고"
        ] if c in reco.columns]
        st.dataframe(reco[out_cols], use_container_width=True, hide_index=True)

# =========================================================
# 8. 원천데이터 탐색 / 다운로드
# =========================================================
elif menu == "8. 원천데이터 탐색 / 다운로드":
    st.header("8. 원천데이터 탐색 / 다운로드")
    info_box(
        "핵심 해석 원칙",
        "앱은 <b>최종 확정본 엑셀</b>만을 데이터 원천으로 사용한다. "
        "원점수·상대위험지수·우선관리지표가 어떤 시트에서 왔는지 직접 확인할 수 있도록 전체 시트를 탐색/다운로드한다."
    )

    sheet_names = list(data["sheets"].keys())
    selected_sheet = st.selectbox("시트 선택", sheet_names)
    sheet_df = clean_columns(data["sheets"][selected_sheet].copy())
    sheet_df = ensure_ym_column(sheet_df)

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

# =========================================================
# 9. 데이터 검증 / 방법론
# =========================================================
elif menu == "9. 데이터 검증 / 방법론":
    st.header("9. 데이터 검증 / 방법론")
    info_box(
        "핵심 해석 원칙",
        "이 메뉴는 <b>왜 이 값이 나왔는지</b>를 추적하기 위한 검증 화면이다. "
        "특히 <b>최종위험점수(원점수) → 상대위험지수(정규화값)</b> 변환과 체인별 EWM 가중치를 투명하게 확인한다."
    )

    tab1, tab2, tab3, tab4 = st.tabs(["방법론", "체인별 가중치", "정규화 검증", "선택월 계산 감사"])

    with tab1:
        st.dataframe(method, use_container_width=True, hide_index=True)
        st.markdown(
            """
            - 가격리스크: 체인별 관련 시장지표(환율/납/리튬/니켈) EWM 결합  
            - 수급리스크: 경보점수기초 + 국가보정합계 EWM 결합  
            - 물류리스크: `GSCPI_Norm`  
            - 정책이벤트리스크: `TPU_INDEX_NORM`  
            - 최종위험점수(원점수): 위 4개 축 EWM 결합  
            - 상대위험지수: 체인별 원점수를 0~100으로 min-max 정규화
            """
        )

    with tab2:
        sel_stage = st.selectbox("가중치 단계 선택", entropy["단계"].dropna().unique().tolist())
        w = entropy[(entropy["단계"] == sel_stage)].copy()
        st.dataframe(w, use_container_width=True, hide_index=True)
        wsum = w.groupby(["단계", "체인구분"], as_index=False)["가중치"].sum()
        st.markdown("##### 가중치 합 점검")
        st.dataframe(wsum, use_container_width=True, hide_index=True)

    with tab3:
        st.dataframe(norm_check, use_container_width=True, hide_index=True)

    with tab4:
        prow = panel_month[panel_month["체인구분"] == selected_chain]
        arow = alert_month[alert_month["체인구분"] == selected_chain]
        crow = compare_chain.copy()
        if prow.empty or arow.empty or crow.empty:
            st.warning("선택월 감사에 필요한 데이터가 부족합니다.")
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
                ["원점수 기반 재계산 상대위험지수", calc_rel],
                ["시트상 상대위험지수", arow["상대위험지수(정규화값)"]],
                ["상대위험지수 Q25", crow["상대위험지수_Q25"]],
                ["상대위험지수 Q50", crow["상대위험지수_Q50"]],
                ["상대위험지수 Q75", crow["상대위험지수_Q75"]],
                ["원점수 Q25", crow["최종위험점수(원점수)_Q25"]],
                ["원점수 Q50", crow["최종위험점수(원점수)_Q50"]],
                ["원점수 Q75", crow["최종위험점수(원점수)_Q75"]],
                ["우선관리대상", arow["상대적_우선관리대상"]],
                ["Q75 초과폭", arow["Q75_초과폭"]],
                ["Q25 미달폭", arow["Q25_미달폭"]],
            ], columns=["항목", "값"])
            st.dataframe(audit_df, use_container_width=True, hide_index=True)

            st.markdown(
                f"""
                **해석**  
                - {selected_chain}의 {selected_month} 원점수는 **{fmt_num(raw_val)}**이고, 체인 내부 최소/최대는 **{fmt_num(raw_min)} / {fmt_num(raw_max)}**이다.  
                - 이 원점수를 기준으로 재계산한 상대위험지수는 **{fmt_num(calc_rel)}**이며, 시트상 값은 **{fmt_num(arow["상대위험지수(정규화값)"])}**이다.  
                - 따라서 앱은 정규화된 상대지표만 단독 제시하지 않고, 항상 원점수와 분위 경계값을 함께 보여주도록 설계했다.
                """
            )
