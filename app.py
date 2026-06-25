import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="배터리 공급망 조기경보·대응 의사결정 시스템", layout="wide")

# =========================
# 공통 유틸
# =========================
def parse_ym_value(v):
    if pd.isna(v):
        return None
    s = str(v).strip()

    if "-" in s:
        p = s.split("-")
        if len(p) >= 2 and p[0].isdigit() and p[1].isdigit():
            y, m = int(p[0]), int(p[1])
            if 1 <= m <= 12:
                return y * 100 + m

    if "." in s:
        left, right = s.split(".", 1)
        left = "".join(ch for ch in left if ch.isdigit())
        right = right.strip()
        if len(left) == 4:
            y = int(left)

            # 원본 특수 규칙: 2021.1 = 2021-10
            if right == "1":
                return y * 100 + 10

            if right.isdigit():
                m = int(right)
                if 1 <= m <= 12:
                    return y * 100 + m

    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 6:
        y, m = int(digits[:4]), int(digits[4:])
        if 1 <= m <= 12:
            return y * 100 + m

    if len(digits) == 5:
        y, m = int(digits[:4]), int(digits[4:])
        if 1 <= m <= 9:
            return y * 100 + m

    return None


def ym_to_label(ym):
    if pd.isna(ym) or ym is None:
        return None
    ym = int(ym)
    return f"{ym // 100}-{ym % 100:02d}"


def fmt_num(x, d=2):
    if pd.isna(x):
        return "-"
    return f"{float(x):,.{d}f}"


def minmax100(series):
    s = pd.to_numeric(series, errors="coerce")
    mn, mx = s.min(), s.max()
    if pd.isna(mn) or pd.isna(mx) or mn == mx:
        return pd.Series(np.where(s.notna(), 50.0, np.nan), index=s.index)
    return (s - mn) / (mx - mn) * 100


def safe_numeric(df, cols):
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def standardize_columns(df, sheet_name=None):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_map = {
        "GSCPI_NORM": "GSCPI_Norm",
        "이벤트보정": "TPU_Norm",
    }
    for old, new in rename_map.items():
        if old in df.columns and new not in df.columns:
            df = df.rename(columns={old: new})

    return df


def add_month_std(df):
    df = df.copy()
    if "연월" in df.columns:
        df["연월_sort"] = df["연월"].apply(parse_ym_value)
        df["연월_표시"] = df["연월_sort"].apply(ym_to_label)
        year_series = pd.to_numeric(df["연월_sort"], errors="coerce")
        df["연도_std"] = np.where(year_series.notna(), (year_series // 100).astype("Int64"), pd.NA)
    return df


def dedup_month_chain(df):
    if df is None or df.empty:
        return df
    if {"연월_표시", "체인구분"}.issubset(df.columns):
        return (
            df.sort_values(["연월_sort", "체인구분"])
            .drop_duplicates(["연월_표시", "체인구분"], keep="first")
            .reset_index(drop=True)
        )
    return df


def get_latest_month(months):
    return months[-1] if months else None


def calc_lag_corr(df, x_col, y_col, max_lag=3):
    temp = df[[x_col, y_col]].copy().dropna()
    out = []
    for lag in range(0, max_lag + 1):
        xx = temp[x_col]
        yy = temp[y_col].shift(-lag)
        z = pd.concat([xx, yy], axis=1).dropna()
        corr = z.iloc[:, 0].corr(z.iloc[:, 1]) if len(z) >= 6 else np.nan
        out.append({"lag_month": lag, "corr": corr})
    return pd.DataFrame(out)


def get_entropy_weight(entropy_df, chain, var_name):
    if entropy_df is None or entropy_df.empty:
        return np.nan
    temp = entropy_df.copy()
    for c in ["체인구분", "구분", "변수명"]:
        if c in temp.columns:
            temp[c] = temp[c].astype(str)
    rr = temp[
        (temp["체인구분"] == str(chain)) &
        (temp["구분"] == "내부가중치") &
        (temp["변수명"] == str(var_name))
    ]
    if rr.empty:
        return np.nan
    return pd.to_numeric(rr.iloc[0]["가중치(%)"], errors="coerce")


# =========================
# 데이터 로드
# =========================
@st.cache_data
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}

    for sh in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sh)
        df = standardize_columns(df, sh)
        df = add_month_std(df)
        sheets[sh] = df

    numeric_targets = {
        "PANEL_MONTHLY": [
            "총수입금액", "총수입물량", "평균수입단가", "상위1국의존도", "상위3국집중도",
            "HHI", "수입국수", "CV", "경보점수기초", "국가보정합계",
            "환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화",
            "GSCPI", "GSCPI_Norm", "원천_TPU_INDEX", "TPU_Norm",
            "환율정규화_5pt", "납가격정규화_5pt", "리튬가격정규화_5pt", "니켈가격정규화_5pt",
            "HHI_5pt", "국가보정합계_5pt", "GSCPI_Norm_5pt", "TPU_Norm_5pt",
            "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
            "최종위험점수_raw", "최종위험점수"
        ],
        "ALERT_RESULT": [
            "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수",
            "최종위험점수", "fta_ratio", "대체조달가능성_점수"
        ],
        "COUNTRY_MONTHLY": [
            "국가수입중량", "국가수입금액", "국가별수입비중", "지역권별수입비중",
            "기본평가점수", "총보정점수", "최종보정점수"
        ],
        "NOMALIZATION_CHECK": [
            "환율", "월평균납가격", "월평균리튬가격", "월평균니켈가격",
            "환율정규화", "납가격정규화", "리튬가격정규화", "니켈가격정규화",
            "GSCPI", "GSCPI_Norm", "원천_TPU_INDEX", "TPU_Norm"
        ],
        "RISK_MASTER": [
            "WGI_Risk_Score_Used", "GPI_Risk_Score", "UCDP_Risk_Score",
            "ACLED_Risk_Score", "Composite_Risk_Score"
        ],
        "Country_Risk_Yearly": [
            "기본위험점수", "분쟁건수", "고유분쟁수", "최고분쟁강도",
            "평균분쟁강도", "개별분쟁점수합"
        ]
    }

    for sh, cols in numeric_targets.items():
        if sh in sheets:
            sheets[sh] = safe_numeric(sheets[sh], cols)

    if "PANEL_MONTHLY" in sheets:
        sheets["PANEL_MONTHLY"] = dedup_month_chain(sheets["PANEL_MONTHLY"])
    if "ALERT_RESULT" in sheets:
        sheets["ALERT_RESULT"] = dedup_month_chain(sheets["ALERT_RESULT"])

    if "체인별 비교표" in sheets:
        df = sheets["체인별 비교표"].copy()
        if "최고_위험_도달_시점" in df.columns:
            df["최고_위험_도달_시점_표시"] = df["최고_위험_도달_시점"].apply(
                lambda x: x.strftime("%Y-%m") if isinstance(x, pd.Timestamp)
                else ym_to_label(parse_ym_value(x)) if pd.notna(x) else None
            )
        sheets["체인별 비교표"] = df

    if "RISK_MASTER" in sheets:
        rm = sheets["RISK_MASTER"].copy()
        rm["Country_Code"] = rm["Country_Code"].astype(str).str.strip().str.upper()
        rm["Year"] = pd.to_numeric(rm["Year"], errors="coerce").astype("Int64")
        sheets["RISK_MASTER"] = rm

    if "COUNTRY_MONTHLY" in sheets:
        cm = sheets["COUNTRY_MONTHLY"].copy()
        cm["국가코드"] = cm["국가코드"].astype(str).str.strip().str.upper()
        if "FTA여부" in cm.columns:
            cm["FTA여부"] = cm["FTA여부"].astype(str).str.upper().str.strip()
        sheets["COUNTRY_MONTHLY"] = cm

    if "Country_Risk_Yearly" in sheets:
        cy = sheets["Country_Risk_Yearly"].copy()
        if "연도" in cy.columns:
            cy["연도"] = pd.to_numeric(cy["연도"], errors="coerce").astype("Int64")
        sheets["Country_Risk_Yearly"] = cy

    return sheets


# =========================
# 분석 유틸
# =========================
def make_lead_table(panel_sub, chain_name):
    candidates = [
        ("환율정규화", "가격리스크점수"),
        ("GSCPI_Norm", "물류리스크점수"),
        ("TPU_Norm", "최종위험점수"),
    ]

    if chain_name == "납산배터리군":
        candidates.append(("납가격정규화", "가격리스크점수"))
    else:
        candidates.extend([
            ("리튬가격정규화", "가격리스크점수"),
            ("니켈가격정규화", "가격리스크점수"),
        ])

    rows = []
    for x_col, y_col in candidates:
        if x_col in panel_sub.columns and y_col in panel_sub.columns:
            corr_df = calc_lag_corr(panel_sub, x_col, y_col, 3)
            valid = corr_df.dropna()
            if not valid.empty:
                best = valid.loc[valid["corr"].abs().idxmax()]
                rows.append({
                    "선행변수": x_col,
                    "반응변수": y_col,
                    "최대연동 시차(개월)": int(best["lag_month"]),
                    "상관계수": round(float(best["corr"]), 4),
                    "해석": f"{x_col} 변화 후 약 {int(best['lag_month'])}개월 내 {y_col} 연동 가능성"
                })
    return pd.DataFrame(rows)


def recommend_countries(country_df, risk_master_df, chain, month, top_n=7):
    if country_df is None or country_df.empty:
        return pd.DataFrame()

    sub = country_df[
        (country_df["체인구분"] == chain) &
        (country_df["연월_표시"] == month)
    ].copy()

    if sub.empty:
        return pd.DataFrame()

    agg = sub.groupby(["국가코드", "국가명"], as_index=False).agg({
        "국가수입금액": "sum",
        "국가수입중량": "sum",
        "국가별수입비중": "mean",
        "최종보정점수": "mean",
        "FTA여부": "max",
    })

    if risk_master_df is not None and not risk_master_df.empty:
        year = int(sub["연도_std"].dropna().iloc[0])
        rm = risk_master_df[risk_master_df["Year"] == year][
            ["Country_Code", "Composite_Risk_Score", "Risk_Flag", "Country"]
        ].copy()
        agg = agg.merge(rm, left_on="국가코드", right_on="Country_Code", how="left")

    agg["risk_component"] = 100 - pd.to_numeric(agg.get("최종보정점수"), errors="coerce").fillna(50)
    agg["fta_component"] = np.where(agg["FTA여부"].astype(str).eq("Y"), 100, 40)
    agg["supply_component"] = minmax100(agg["국가수입금액"]).fillna(50)
    agg["div_component"] = 100 - pd.to_numeric(agg["국가별수입비중"], errors="coerce").fillna(0)

    if "Composite_Risk_Score" in agg.columns:
        agg["external_risk_component"] = 100 - pd.to_numeric(
            agg["Composite_Risk_Score"], errors="coerce"
        ).fillna(50)
        agg["대체국추천점수"] = (
            0.25 * agg["risk_component"] +
            0.20 * agg["fta_component"] +
            0.15 * agg["supply_component"] +
            0.15 * agg["div_component"] +
            0.25 * agg["external_risk_component"]
        ).round(2)
    else:
        agg["대체국추천점수"] = (
            0.35 * agg["risk_component"] +
            0.25 * agg["fta_component"] +
            0.20 * agg["supply_component"] +
            0.20 * agg["div_component"]
        ).round(2)

    agg["추천사유"] = (
        "내부위험 낮음 / " +
        np.where(agg["FTA여부"].astype(str).eq("Y"), "FTA 유리 / ", "FTA 비우위 / ") +
        np.where(agg.get("Risk_Flag").notna(), "대외위험 근거 반영 / ", "") +
        "집중도 완화 가능"
    )

    return (
        agg.sort_values(["대체국추천점수", "국가수입금액"], ascending=[False, False])
        .head(top_n)
        .reset_index(drop=True)
    )


def simulate_scenario(panel_row, alert_row, top1_drop=10, fta_gain=10, alt_gain=10):
    p = panel_row.iloc[0]
    a = alert_row.iloc[0]

    current_top1 = float(p.get("상위1국의존도", np.nan))
    current_hhi = float(p.get("HHI", np.nan))
    current_score = float(a.get("최종위험점수", np.nan))
    current_fta = float(a.get("fta_ratio", np.nan))
    current_alt = float(a.get("대체조달가능성_점수", np.nan))

    new_top1 = max(current_top1 - top1_drop, 0) if pd.notna(current_top1) else np.nan
    new_hhi = max(current_hhi * (1 - 0.6 * top1_drop / 100), 0) if pd.notna(current_hhi) else np.nan
    new_fta = min(current_fta + fta_gain, 100) if pd.notna(current_fta) else np.nan
    new_alt = min(current_alt + alt_gain, 100) if pd.notna(current_alt) else np.nan

    score_reduction = 0.25 * top1_drop + 0.08 * fta_gain + 0.12 * alt_gain
    new_score = max(current_score - score_reduction, 0) if pd.notna(current_score) else np.nan

    return {
        "현재 상위1국의존도": current_top1,
        "시뮬레이션 상위1국의존도": round(new_top1, 2) if pd.notna(new_top1) else np.nan,
        "현재 HHI": current_hhi,
        "시뮬레이션 HHI": round(new_hhi, 2) if pd.notna(new_hhi) else np.nan,
        "현재 FTA 활용비중": current_fta,
        "시뮬레이션 FTA 활용비중": round(new_fta, 2) if pd.notna(new_fta) else np.nan,
        "현재 대체조달가능성점수": current_alt,
        "시뮬레이션 대체조달가능성점수": round(new_alt, 2) if pd.notna(new_alt) else np.nan,
        "현재 최종위험점수": current_score,
        "시뮬레이션 최종위험점수": round(new_score, 2) if pd.notna(new_score) else np.nan,
        "예상 감소폭": round(current_score - new_score, 2)
        if pd.notna(current_score) and pd.notna(new_score) else np.nan
    }


def build_country_risk_profile(country_monthly_df, risk_master_df, country_risk_yearly_df, conflict_df, chain, month, country_code):
    sub = country_monthly_df[
        (country_monthly_df["체인구분"] == chain) &
        (country_monthly_df["연월_표시"] == month) &
        (country_monthly_df["국가코드"] == country_code)
    ].copy()

    if sub.empty:
        return None, pd.DataFrame(), pd.DataFrame()

    year = int(sub["연도_std"].dropna().iloc[0])

    base = (
        sub.groupby(["국가코드", "국가명"], as_index=False)
        .agg({
            "국가수입금액": "sum",
            "국가별수입비중": "mean",
            "기본평가점수": "mean",
            "총보정점수": "mean",
            "최종보정점수": "mean",
            "FTA여부": "max",
        })
        .iloc[0]
        .to_dict()
    )

    rm = pd.DataFrame()
    if risk_master_df is not None and not risk_master_df.empty:
        rm = risk_master_df[
            (risk_master_df["Country_Code"] == country_code) &
            (risk_master_df["Year"] == year)
        ].copy()

    cf = pd.DataFrame()
    if conflict_df is not None and not conflict_df.empty and not rm.empty:
        conflict_df = conflict_df.copy()
        if "연도_std" not in conflict_df.columns:
            conflict_df = add_month_std(conflict_df)
        en_name = rm["Country"].iloc[0]
        cf = conflict_df[
            (conflict_df["연도_std"] == year) &
            (conflict_df["국가명"].astype(str) == str(en_name))
        ].copy()

    return base, rm, cf


# =========================
# UI 시작
# =========================
st.title("🔎 배터리 공급망 조기경보·대응 의사결정 시스템")
st.caption("위험 진단 · 원인 설명 · 선행 신호 탐지 · 대체국 추천 · 대응 시뮬레이션")

uploaded_file = st.file_uploader("최종 엑셀 파일 업로드", type=["xlsx"])

if uploaded_file is None:
    st.info("최종 엑셀 파일을 업로드해 주세요.")
    st.stop()

sheets = load_data(uploaded_file)

panel = sheets.get("PANEL_MONTHLY")
alert = sheets.get("ALERT_RESULT")
compare_df = sheets.get("체인별 비교표")
entropy = sheets.get("ENTROPY_WEIGHT")
norm_check = sheets.get("NOMALIZATION_CHECK")
country = sheets.get("COUNTRY_MONTHLY")
tpu = sheets.get("TPU_INDEX")
risk_master = sheets.get("RISK_MASTER")
country_risk_yearly = sheets.get("Country_Risk_Yearly")
conflict = sheets.get("Conflict_Detail")
hs = sheets.get("HS_MONTHLY_SUMMARY")

required = {
    "PANEL_MONTHLY": panel,
    "ALERT_RESULT": alert,
    "체인별 비교표": compare_df,
    "ENTROPY_WEIGHT": entropy,
    "NOMALIZATION_CHECK": norm_check,
}
missing = [k for k, v in required.items() if v is None]
if missing:
    st.error(f"필수 시트 누락: {missing}")
    st.stop()

months = sorted(panel["연월_표시"].dropna().unique().tolist())
chains = sorted(panel["체인구분"].dropna().unique().tolist())
latest_month = get_latest_month(months)

st.sidebar.title("분석 메뉴")
menu = st.sidebar.radio(
    "선택",
    [
        "1. 종합 상황판",
        "2. 체인별 심층 분석",
        "3. 충격 원인 추적",
        "4. 선행 신호 탐지",
        "5. 기업 대응 시뮬레이터",
        "6. 대체국 추천 시스템",
        "7. 데이터 검증 / 방법론",
    ],
)

# =========================
# 1. 종합 상황판
# =========================
if menu == "1. 종합 상황판":
    st.header("1. 종합 상황판")

    curr = alert[alert["연월_표시"] == latest_month].copy()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기준월", latest_month)
    c2.metric("분석 체인 수", len(curr))
    c3.metric(
        "우선관리대상 수",
        int((curr["상대적_우선관리대상"] == "Y").sum()) if "상대적_우선관리대상" in curr.columns else 0,
    )
    c4.metric("평균 최종위험점수", fmt_num(curr["최종위험점수"].mean(), 2))

    fig = px.bar(
        curr.sort_values("최종위험점수", ascending=False),
        x="체인구분",
        y="최종위험점수",
        color="최종경보등급",
        text_auto=".2f",
        title=f"{latest_month} 체인별 최종위험점수",
    )
    st.plotly_chart(fig, use_container_width=True)

    cols = [
        c
        for c in [
            "체인구분", "최종위험점수", "최종경보등급",
            "보정사유", "대체조달가능성", "상대적_우선관리대상", "비고"
        ]
        if c in curr.columns
    ]
    st.dataframe(curr[cols], use_container_width=True)

    if compare_df is not None and not compare_df.empty:
        st.markdown("### 체인별 장기 비교")
        show_cols = [
            c
            for c in [
                "체인구분", "평균_최종위험점수", "최고 위험점수",
                "최고_위험_도달_시점_표시", "평균_FTA_활용비중",
                "평균_대체조달가능성점수", "우선관리대상 비중 (%)", "종합_시사점"
            ]
            if c in compare_df.columns
        ]
        st.dataframe(compare_df[show_cols], use_container_width=True)

# =========================
# 2. 체인별 심층 분석
# =========================
elif menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")

    chain = st.selectbox("체인 선택", chains)
    p = panel[panel["체인구분"] == chain].sort_values("연월_sort").copy()
    a = alert[alert["체인구분"] == chain].sort_values("연월_sort").copy()
    latest = a.iloc[-1]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("최신 최종위험점수", fmt_num(latest["최종위험점수"], 2))
    k2.metric("최신 경보등급", latest["최종경보등급"])
    k3.metric("대체조달가능성", latest["대체조달가능성"])
    k4.metric("FTA 활용비중", fmt_num(latest["fta_ratio"], 2))

    fig1 = px.line(a, x="연월_표시", y="최종위험점수", markers=True, title=f"{chain} 최종위험점수 추이")
    st.plotly_chart(fig1, use_container_width=True)

    vars1 = [c for c in ["HHI", "상위1국의존도", "국가보정합계", "CV"] if c in p.columns]
    fig2 = px.line(p, x="연월_표시", y=vars1, title=f"{chain} 구조 취약도 추이")
    st.plotly_chart(fig2, use_container_width=True)

    vars2 = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in p.columns]
    fig3 = px.line(p, x="연월_표시", y=vars2, title=f"{chain} 4대 리스크 추이")
    st.plotly_chart(fig3, use_container_width=True)

# =========================
# 3. 충격 원인 추적
# =========================
elif menu == "3. 충격 원인 추적":
    st.header("3. 충격 원인 추적")

    c1, c2 = st.columns(2)
    chain = c1.selectbox("체인", chains)
    month = c2.selectbox("연월", months, index=len(months) - 1)

    p = panel[(panel["체인구분"] == chain) & (panel["연월_표시"] == month)].copy()
    a = alert[(alert["체인구분"] == chain) & (alert["연월_표시"] == month)].copy()

    if p.empty or a.empty:
        st.warning("선택 조건 데이터가 없습니다.")
        st.stop()

    p0 = p.iloc[0]
    a0 = a.iloc[0]

    st.markdown("### 현재월 카테고리 기여도")
    risk_df = pd.DataFrame(
        {
            "리스크유형": ["가격", "수급", "물류", "정책이벤트"],
            "점수": [
                a0.get("가격리스크점수"),
                a0.get("수급리스크점수"),
                a0.get("물류리스크점수"),
                a0.get("정책이벤트리스크점수"),
            ],
        }
    ).sort_values("점수", ascending=False)

    fig = px.bar(risk_df, x="리스크유형", y="점수", color="리스크유형", text_auto=".2f")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 가격 내부요인 분해")
    if chain == "납산배터리군":
        detail = pd.DataFrame(
            {
                "요인": ["환율정규화", "납가격정규화"],
                "정규화값": [p0.get("환율정규화"), p0.get("납가격정규화")],
                "5점값": [p0.get("환율정규화_5pt"), p0.get("납가격정규화_5pt")],
                "가중치(%)": [
                    get_entropy_weight(entropy, chain, "환율정규화"),
                    get_entropy_weight(entropy, chain, "납가격정규화"),
                ],
            }
        )
    else:
        detail = pd.DataFrame(
            {
                "요인": ["환율정규화", "리튬가격정규화", "니켈가격정규화"],
                "정규화값": [
                    p0.get("환율정규화"),
                    p0.get("리튬가격정규화"),
                    p0.get("니켈가격정규화"),
                ],
                "5점값": [
                    p0.get("환율정규화_5pt"),
                    p0.get("리튬가격정규화_5pt"),
                    p0.get("니켈가격정규화_5pt"),
                ],
                "가중치(%)": [
                    get_entropy_weight(entropy, chain, "환율정규화"),
                    get_entropy_weight(entropy, chain, "리튬가격정규화"),
                    get_entropy_weight(entropy, chain, "니켈가격정규화"),
                ],
            }
        )
    detail["기여도"] = detail["5점값"] * detail["가중치(%)"] / 100
    st.dataframe(detail, use_container_width=True)

    st.markdown("### 정책 이벤트 서사배경")
    if tpu is not None and "서사배경" in tpu.columns:
        tr = tpu[tpu["연월_표시"] == month]
        if not tr.empty:
            st.info(str(tr.iloc[0]["서사배경"]))
        else:
            st.write("해당 월 서사배경 정보가 없습니다.")
    else:
        st.write("TPU_INDEX 서사배경 정보가 없습니다.")

    st.markdown("### 공급국 리스크 근거 연결")
    if country is not None and not country.empty:
        csub = country[(country["체인구분"] == chain) & (country["연월_표시"] == month)].copy()
        if not csub.empty:
            csub = csub.sort_values("국가별수입비중", ascending=False)
            options = csub["국가코드"].astype(str) + " | " + csub["국가명"].astype(str)
            selected = st.selectbox("국가 선택", options.tolist())
            code = selected.split(" | ")[0]

            base, rm, cf = build_country_risk_profile(
                country, risk_master, country_risk_yearly, conflict, chain, month, code
            )

            if base:
                b1, b2, b3, b4 = st.columns(4)
                b1.metric("국가별수입비중", fmt_num(base.get("국가별수입비중"), 2))
                b2.metric("기본평가점수", fmt_num(base.get("기본평가점수"), 2))
                b3.metric("총보정점수", fmt_num(base.get("총보정점수"), 2))
                b4.metric("최종보정점수", fmt_num(base.get("최종보정점수"), 2))

            if rm is not None and not rm.empty:
                st.markdown("#### 국가 리스크 근거")
                show_rm = [
                    c
                    for c in [
                        "Country", "Risk_Flag", "Composite_Risk_Score",
                        "WGI_Risk_Score_Used", "GPI_Risk_Score",
                        "UCDP_Risk_Score", "ACLED_Risk_Score"
                    ]
                    if c in rm.columns
                ]
                st.dataframe(rm[show_rm], use_container_width=True)

            if cf is not None and not cf.empty:
                st.markdown("#### 최근 분쟁/갈등 근거")
                show_cf = [
                    c
                    for c in ["국가명", "분쟁유형", "분쟁강도명", "불일치유형명", "시작일", "종료일"]
                    if c in cf.columns
                ]
                st.dataframe(cf[show_cf].head(10), use_container_width=True)

# =========================
# 4. 선행 신호 탐지
# =========================
elif menu == "4. 선행 신호 탐지":
    st.header("4. 선행 신호 탐지")

    chain = st.selectbox("체인 선택", chains, key="lead")
    p = panel[panel["체인구분"] == chain].sort_values("연월_sort").copy()
    lead = make_lead_table(p, chain)

    if lead.empty:
        st.warning("분석 가능한 선행신호가 부족합니다.")
    else:
        st.dataframe(lead, use_container_width=True)
        xcol = st.selectbox("선행변수", lead["선행변수"].tolist())
        ycol = lead[lead["선행변수"] == xcol].iloc[0]["반응변수"]

        corr_df = calc_lag_corr(p, xcol, ycol, max_lag=3)
        fig = px.bar(corr_df, x="lag_month", y="corr", text_auto=".3f", title=f"{xcol} → {ycol} 시차별 상관")
        st.plotly_chart(fig, use_container_width=True)
        st.caption("해석 주의: 이는 인과 확정이 아니라 향후 1~3개월 선제 모니터링 신호입니다.")

# =========================
# 5. 기업 대응 시뮬레이터
# =========================
elif menu == "5. 기업 대응 시뮬레이터":
    st.header("5. 기업 대응 시뮬레이터")

    c1, c2 = st.columns(2)
    chain = c1.selectbox("체인", chains, key="sim_chain")
    month = c2.selectbox("연월", months, index=len(months) - 1, key="sim_month")

    p = panel[(panel["체인구분"] == chain) & (panel["연월_표시"] == month)]
    a = alert[(alert["체인구분"] == chain) & (alert["연월_표시"] == month)]

    if p.empty or a.empty:
        st.warning("선택 조건 데이터가 없습니다.")
        st.stop()

    s1 = st.slider("상위1국의존도 감소폭(%p)", 0, 30, 10)
    s2 = st.slider("FTA 활용비중 개선폭(%p)", 0, 30, 10)
    s3 = st.slider("대체조달가능성 개선폭(점수)", 0, 30, 10)

    sim = simulate_scenario(p, a, s1, s2, s3)

    m1, m2, m3 = st.columns(3)
    m1.metric("현재 최종위험점수", fmt_num(sim["현재 최종위험점수"], 2))
    m2.metric("시뮬레이션 최종위험점수", fmt_num(sim["시뮬레이션 최종위험점수"], 2))
    m3.metric("예상 감소폭", fmt_num(sim["예상 감소폭"], 2))

    sim_df = pd.DataFrame(
        {
            "지표": ["상위1국의존도", "HHI", "FTA 활용비중", "대체조달가능성점수", "최종위험점수"],
            "현재": [
                sim["현재 상위1국의존도"], sim["현재 HHI"], sim["현재 FTA 활용비중"],
                sim["현재 대체조달가능성점수"], sim["현재 최종위험점수"]
            ],
            "시뮬레이션": [
                sim["시뮬레이션 상위1국의존도"], sim["시뮬레이션 HHI"], sim["시뮬레이션 FTA 활용비중"],
                sim["시뮬레이션 대체조달가능성점수"], sim["시뮬레이션 최종위험점수"]
            ],
        }
    )
    st.dataframe(sim_df, use_container_width=True)

# =========================
# 6. 대체국 추천 시스템
# =========================
elif menu == "6. 대체국 추천 시스템":
    st.header("6. 대체국 추천 시스템")

    c1, c2 = st.columns(2)
    chain = c1.selectbox("체인", chains, key="alt_chain")
    month = c2.selectbox("연월", months, index=len(months) - 1, key="alt_month")

    rec = recommend_countries(country, risk_master, chain, month, top_n=7)

    if rec.empty:
        st.warning("추천 가능한 국가 데이터가 없습니다.")
    else:
        show_cols = [
            c
            for c in [
                "국가명", "국가코드", "국가수입금액", "국가별수입비중", "최종보정점수",
                "Composite_Risk_Score", "Risk_Flag", "FTA여부", "대체국추천점수", "추천사유"
            ]
            if c in rec.columns
        ]
        st.dataframe(rec[show_cols], use_container_width=True)

        fig = px.bar(
            rec.sort_values("대체국추천점수"),
            x="대체국추천점수",
            y="국가명",
            orientation="h",
            text_auto=".2f",
            title=f"{chain} / {month} 대체국 추천 우선순위",
        )
        st.plotly_chart(fig, use_container_width=True)

# =========================
# 7. 데이터 검증 / 방법론
# =========================
elif menu == "7. 데이터 검증 / 방법론":
    st.header("7. 데이터 검증 / 방법론")

    targets = {
        "PANEL_MONTHLY": panel,
        "ALERT_RESULT": alert,
        "NOMALIZATION_CHECK": norm_check,
        "COUNTRY_MONTHLY": country,
        "TPU_INDEX": tpu,
        "HS_MONTHLY_SUMMARY": hs,
    }

    rows = []
    for name, df in targets.items():
        if df is None:
            continue
        rows.append(
            {
                "시트명": name,
                "행수": len(df),
                "연월표준화결측": int(df["연월_표시"].isna().sum()) if "연월_표시" in df.columns else None,
                "중복키수": int(df.duplicated(["연월_표시", "체인구분"]).sum())
                if {"연월_표시", "체인구분"}.issubset(df.columns) else None,
            }
        )

    st.dataframe(pd.DataFrame(rows), use_container_width=True)

    st.markdown("### 표준화/예외처리 반영사항")
    st.write("- `2021.01`, `2021-01`, `202101`을 `2021-01`로 통일")
    st.write("- 원본 특수값 `2021.1`은 `2021-10`으로 처리")
    st.write("- `GSCPI_NORM → GSCPI_Norm`, `이벤트보정 → TPU_Norm` 자동 통일")
    st.write("- `ALERT_RESULT`의 float형 연월, `체인별 비교표`의 Timestamp형 연월도 화면에서는 통일 표기")

    st.markdown("### 주요 산식")
    st.code(
        "CV = 국가별 월수입금액 population std / 평균\n"
        "경보점수기초 = (HHI/100 × 0.4) + (상위1국의존도 × 0.3) + (CV × 0.3)\n"
        "대체조달가능성_점수 = 0.35*수입국수_norm + 0.30*(100-상위1국의존도) + 0.20*(100-HHI_norm) + 0.15*fta_ratio"
    )
