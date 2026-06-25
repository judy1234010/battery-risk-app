import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(
    page_title="배터리 공급망 조기경보·대응 의사결정 시스템",
    layout="wide"
)

# =========================================================
# 0. 스타일
# =========================================================
st.markdown("""
<style>
.small-note {font-size: 0.88rem; color: #B8C0CC;}
.section-gap {margin-top: 1.2rem; margin-bottom: 0.5rem;}
.badge-red {
    display:inline-block; padding:0.18rem 0.55rem; border-radius:0.5rem;
    background:#5c1f1f; color:#ffd7d7; font-weight:600;
}
.badge-gray {
    display:inline-block; padding:0.18rem 0.55rem; border-radius:0.5rem;
    background:#2f3640; color:#d8dee9; font-weight:600;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# 1. 공통 유틸
# =========================================================
def parse_ym_value(v):
    if pd.isna(v):
        return None

    if isinstance(v, pd.Timestamp):
        return v.year * 100 + v.month

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

            # 원본 특수 규칙
            # 2021.1 = 2021-10
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


def safe_numeric(df, cols):
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def minmax100(series):
    s = pd.to_numeric(series, errors="coerce")
    mn, mx = s.min(), s.max()
    if pd.isna(mn) or pd.isna(mx) or mn == mx:
        return pd.Series(np.where(s.notna(), 50.0, np.nan), index=s.index)
    return (s - mn) / (mx - mn) * 100


def normalize_0_100(series):
    s = pd.to_numeric(series, errors="coerce")
    mn, mx = s.min(), s.max()
    if pd.isna(mn) or pd.isna(mx) or mn == mx:
        return pd.Series(np.where(s.notna(), 50.0, np.nan), index=s.index)
    return (s - mn) / (mx - mn) * 100


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
        s = pd.to_numeric(df["연월_sort"], errors="coerce")
        df["연도_std"] = np.where(s.notna(), (s // 100).astype("Int64"), pd.NA)
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


# =========================================================
# 2. 데이터 로드
# =========================================================
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
            "총수입금액", "총수입물량", "평균수입단가",
            "상위1국의존도", "상위3국집중도", "HHI", "수입국수", "CV",
            "경보점수기초", "국가보정합계",
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
        ],
        "HS_MONTHLY_SUMMARY": [
            "총수입중량", "총수입금액", "평균수입단가", "수입국수",
            "상위공급국비중", "전월대비수입금액증감률", "전월대비수입물량증감률"
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
        if "Country_Code" in rm.columns:
            rm["Country_Code"] = rm["Country_Code"].astype(str).str.strip().str.upper()
        if "Year" in rm.columns:
            rm["Year"] = pd.to_numeric(rm["Year"], errors="coerce").astype("Int64")
        sheets["RISK_MASTER"] = rm

    if "COUNTRY_MONTHLY" in sheets:
        cm = sheets["COUNTRY_MONTHLY"].copy()
        if "국가코드" in cm.columns:
            cm["국가코드"] = cm["국가코드"].astype(str).str.strip().str.upper()
        if "FTA여부" in cm.columns:
            cm["FTA여부"] = cm["FTA여부"].astype(str).str.upper().str.strip()
        sheets["COUNTRY_MONTHLY"] = cm

    if "Country_Risk_Yearly" in sheets:
        cy = sheets["Country_Risk_Yearly"].copy()
        if "연도" in cy.columns:
            cy["연도"] = pd.to_numeric(cy["연도"], errors="coerce").astype("Int64")
        sheets["Country_Risk_Yearly"] = cy

    if "Conflict_Detail" in sheets:
        cd = sheets["Conflict_Detail"].copy()
        if "연도" in cd.columns:
            cd["연도"] = pd.to_numeric(cd["연도"], errors="coerce").astype("Int64")
        sheets["Conflict_Detail"] = cd

    return sheets


# =========================================================
# 3. 선행 신호 탐지
#    x(t) -> y(t+lag), lag=1~3만 사용
# =========================================================
def corr_strength_label(corr):
    if pd.isna(corr):
        return "판단불가"
    a = abs(float(corr))
    if a >= 0.5:
        return "강함"
    if a >= 0.3:
        return "보통"
    return "약함"


def build_lead_interpretation(x_col, y_col, lag, corr):
    if pd.isna(corr):
        return "표본이 충분하지 않아 시차 연동성을 판단하기 어렵습니다."

    direction = "양(+)의" if corr > 0 else "음(-)의"
    strength = corr_strength_label(corr)

    if strength == "약함":
        return (
            f"{x_col}와 {y_col}는 {lag}개월 선행 구간에서 {direction} 상관이 관찰되지만 "
            f"강도가 약해 참고용 보조 신호로 해석하는 것이 적절합니다."
        )

    return (
        f"{x_col}가 약 {lag}개월 선행하는 구간에서 {y_col}와 {direction} 연동성이 관찰됩니다. "
        f"이는 인과 확정이 아니라 향후 {lag}개월 내 선제 모니터링 우선순위를 제시하는 보조 신호입니다."
    )


def calc_forward_lag_corr(df, x_col, y_col, lags=(1, 2, 3)):
    temp = df[[x_col, y_col]].copy()
    temp[x_col] = pd.to_numeric(temp[x_col], errors="coerce")
    temp[y_col] = pd.to_numeric(temp[y_col], errors="coerce")
    temp = temp.dropna()

    rows = []
    for lag in lags:
        z = pd.DataFrame({
            "x": temp[x_col],
            "y_future": temp[y_col].shift(-lag)
        }).dropna()

        corr = z["x"].corr(z["y_future"]) if len(z) >= 6 else np.nan
        rows.append({
            "lag_month": lag,
            "corr": corr,
            "n_obs": len(z)
        })

    return pd.DataFrame(rows)


def get_lead_candidates(chain_name, df):
    base = ["TPU_Norm", "GSCPI_Norm", "국가보정합계", "HHI", "상위1국의존도", "CV", "환율정규화"]

    if chain_name == "납산배터리군":
        base += ["납가격정규화"]
    else:
        base += ["리튬가격정규화", "니켈가격정규화"]

    return [c for c in base if c in df.columns]


def make_lead_table(panel_sub, chain_name, future_target):
    candidates = get_lead_candidates(chain_name, panel_sub)
    rows = []

    for x_col in candidates:
        corr_df = calc_forward_lag_corr(panel_sub, x_col, future_target, lags=(1, 2, 3))
        valid = corr_df.dropna(subset=["corr"])
        if valid.empty:
            continue

        best = valid.loc[valid["corr"].abs().idxmax()]
        strength = corr_strength_label(best["corr"])

        rows.append({
            "선행변수": x_col,
            "미래반응지표": future_target,
            "최대연동 시차(개월)": int(best["lag_month"]),
            "상관계수": round(float(best["corr"]), 4),
            "강도": strength,
            "표본수": int(best["n_obs"]),
            "해석": build_lead_interpretation(x_col, future_target, int(best["lag_month"]), float(best["corr"]))
        })

    if not rows:
        return pd.DataFrame()

    out = pd.DataFrame(rows)
    order_map = {"강함": 2, "보통": 1, "약함": 0, "판단불가": -1}
    out["_ord"] = out["강도"].map(order_map)
    out = out.sort_values(["_ord", "상관계수"], ascending=[False, False]).drop(columns="_ord")
    return out.reset_index(drop=True)


# =========================================================
# 4. 국가 리스크 / 대체국 추천
# =========================================================
def build_country_risk_profile(country_monthly_df, risk_master_df, conflict_df, chain, month, country_code):
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
        en_name = rm["Country"].iloc[0]
        cf = conflict_df[
            (conflict_df["연도"] == year) &
            (conflict_df["국가명"].astype(str) == str(en_name))
        ].copy()

    return base, rm, cf


def build_country_explanation(base, rm):
    if not base:
        return "국가별 리스크 설명 정보를 생성할 수 없습니다."

    parts = []
    country_name = base.get("국가명", "해당 국가")
    share = base.get("국가별수입비중")
    final_score = base.get("최종보정점수")
    fta = base.get("FTA여부", "N")

    if pd.notna(share):
        if share >= 40:
            parts.append("수입 비중이 높아 집중 리스크가 큽니다")
        elif share >= 15:
            parts.append("유의미한 조달 비중을 차지합니다")
        else:
            parts.append("현재 비중은 크지 않지만 분산 후보로 검토 가능합니다")

    if pd.notna(final_score):
        if final_score >= 70:
            parts.append("내부 공급망 보정점수가 높아 구조적 주의가 필요합니다")
        elif final_score >= 40:
            parts.append("내부 위험 수준은 중간 이상입니다")
        else:
            parts.append("내부 위험 수준은 상대적으로 낮은 편입니다")

    if str(fta).upper() == "Y":
        parts.append("FTA 활용 측면에서 유리합니다")
    else:
        parts.append("FTA 측면 우위는 제한적입니다")

    if rm is not None and not rm.empty:
        rr = rm.iloc[0]
        comp = rr.get("Composite_Risk_Score", np.nan)
        flag = rr.get("Risk_Flag", None)

        if pd.notna(comp):
            if comp >= 40:
                parts.append("대외 국가리스크가 비교적 높은 편입니다")
            elif comp >= 20:
                parts.append("대외 국가리스크는 중간 수준입니다")
            else:
                parts.append("대외 국가리스크는 비교적 낮은 편입니다")

        if pd.notna(flag):
            parts.append(f"종합 위험등급은 '{flag}' 수준입니다")

    return f"{country_name}: " + " / ".join(parts) + "."


def recommend_country_groups(country_df, risk_master_df, chain, month):
    if country_df is None or country_df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    sub = country_df[
        (country_df["체인구분"] == chain) &
        (country_df["연월_표시"] == month)
    ].copy()

    if sub.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    agg = sub.groupby(["국가코드", "국가명"], as_index=False).agg({
        "국가수입금액": "sum",
        "국가수입중량": "sum",
        "국가별수입비중": "mean",
        "최종보정점수": "mean",
        "기본평가점수": "mean",
        "FTA여부": "max",
    })

    year = int(sub["연도_std"].dropna().iloc[0])

    if risk_master_df is not None and not risk_master_df.empty:
        rm = risk_master_df[risk_master_df["Year"] == year][
            ["Country_Code", "Country", "Composite_Risk_Score", "Risk_Flag"]
        ].copy()
        agg = agg.merge(rm, left_on="국가코드", right_on="Country_Code", how="left")

    agg["내부안정성"] = 100 - pd.to_numeric(agg["최종보정점수"], errors="coerce").fillna(50)
    agg["대외안정성"] = 100 - pd.to_numeric(agg.get("Composite_Risk_Score"), errors="coerce").fillna(50)
    agg["FTA우위"] = np.where(agg["FTA여부"].astype(str).eq("Y"), 100, 40)
    agg["실행가능성"] = minmax100(agg["국가수입금액"]).fillna(50)
    agg["분산효과"] = 100 - pd.to_numeric(agg["국가별수입비중"], errors="coerce").fillna(0)

    agg["유지추천점수"] = (
        0.30 * agg["내부안정성"] +
        0.25 * agg["대외안정성"] +
        0.20 * agg["FTA우위"] +
        0.25 * agg["실행가능성"]
    ).round(2)

    agg["확대추천점수"] = (
        0.25 * agg["내부안정성"] +
        0.20 * agg["대외안정성"] +
        0.20 * agg["FTA우위"] +
        0.10 * agg["실행가능성"] +
        0.25 * agg["분산효과"]
    ).round(2)

    agg["축소검토점수"] = (
        0.40 * pd.to_numeric(agg["국가별수입비중"], errors="coerce").fillna(0) +
        0.35 * pd.to_numeric(agg["최종보정점수"], errors="coerce").fillna(50) +
        0.25 * pd.to_numeric(agg.get("Composite_Risk_Score"), errors="coerce").fillna(50)
    ).round(2)

    def make_keep_reason(r):
        parts = []
        if r["FTA여부"] == "Y":
            parts.append("FTA 활용 가능")
        if pd.notna(r.get("Composite_Risk_Score")) and r["Composite_Risk_Score"] < 25:
            parts.append("대외 국가리스크 낮음")
        if pd.notna(r["최종보정점수"]) and r["최종보정점수"] < 40:
            parts.append("내부 공급망 위험도 낮음")
        if pd.notna(r["국가수입금액"]) and r["국가수입금액"] > 0:
            parts.append("기존 거래 실적 보유")
        return " / ".join(parts) if parts else "안정성과 실행가능성을 종합 고려할 때 유지 가치가 있음"

    def make_expand_reason(r):
        parts = []
        if r["FTA여부"] == "Y":
            parts.append("FTA 활용 여지")
        if pd.notna(r["국가별수입비중"]) and r["국가별수입비중"] < 15:
            parts.append("현재 비중이 낮아 분산효과 큼")
        if pd.notna(r.get("Composite_Risk_Score")) and r["Composite_Risk_Score"] < 30:
            parts.append("대외 위험 상대적으로 안정적")
        return " / ".join(parts) if parts else "비중 확대 시 집중도 완화 효과 기대"

    def make_reduce_reason(r):
        parts = []
        if pd.notna(r["국가별수입비중"]) and r["국가별수입비중"] >= 20:
            parts.append("현재 비중이 높음")
        if pd.notna(r["최종보정점수"]) and r["최종보정점수"] >= 50:
            parts.append("내부 위험도 높음")
        if pd.notna(r.get("Composite_Risk_Score")) and r["Composite_Risk_Score"] >= 30:
            parts.append("대외 국가리스크 부담")
        return " / ".join(parts) if parts else "집중도 완화 차원에서 축소 검토 가능"

    agg["유지추천사유"] = agg.apply(make_keep_reason, axis=1)
    agg["확대추천사유"] = agg.apply(make_expand_reason, axis=1)
    agg["축소검토사유"] = agg.apply(make_reduce_reason, axis=1)

    keep_df = agg.sort_values(["유지추천점수", "국가수입금액"], ascending=[False, False]).head(5).reset_index(drop=True)
    expand_df = agg[
        pd.to_numeric(agg["국가별수입비중"], errors="coerce").fillna(0) < 25
    ].sort_values(["확대추천점수", "국가수입금액"], ascending=[False, False]).head(5).reset_index(drop=True)
    reduce_df = agg[
        pd.to_numeric(agg["국가별수입비중"], errors="coerce").fillna(0) >= 10
    ].sort_values(["축소검토점수", "국가별수입비중"], ascending=[False, False]).head(5).reset_index(drop=True)

    return keep_df, expand_df, reduce_df


# =========================================================
# 5. 시뮬레이터
# =========================================================
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


# =========================================================
# 6. 검증 로직
# =========================================================
def get_validation_rules():
    return {
        "PANEL_MONTHLY": {
            "key": ["연월_표시", "체인구분"],
            "required": ["연월", "체인구분", "최종위험점수", "최종경보등급"],
            "critical_numeric": ["최종위험점수", "가격리스크점수", "수급리스크점수"]
        },
        "ALERT_RESULT": {
            "key": ["연월_표시", "체인구분"],
            "required": ["연월", "체인구분", "최종위험점수", "최종경보등급", "대체조달가능성"],
            "critical_numeric": ["최종위험점수", "fta_ratio", "대체조달가능성_점수"]
        },
        "NOMALIZATION_CHECK": {
            "key": ["연월_표시"],
            "required": ["연월", "환율정규화", "GSCPI_Norm", "TPU_Norm"],
            "critical_numeric": ["환율정규화", "GSCPI_Norm", "TPU_Norm"]
        },
        "COUNTRY_MONTHLY": {
            "key": ["연월_표시", "체인구분", "국가코드", "HS코드"],
            "required": ["연월", "체인구분", "국가코드", "국가명", "최종보정점수"],
            "critical_numeric": ["국가수입금액", "국가별수입비중", "최종보정점수"]
        },
        "HS_MONTHLY_SUMMARY": {
            "key": ["연월_표시", "체인구분", "HS코드"],
            "required": ["연월", "체인구분", "HS코드", "총수입금액"],
            "critical_numeric": ["총수입금액", "총수입중량", "평균수입단가"]
        },
        "TPU_INDEX": {
            "key": ["연월_표시"],
            "required": ["연월", "원천_TPU_INDEX", "TPU_Norm", "서사배경"],
            "critical_numeric": ["원천_TPU_INDEX", "TPU_Norm"]
        }
    }


def validate_sheet(name, df, rules):
    result = {
        "시트명": name,
        "행수": len(df) if df is not None else None,
        "연월파싱실패수": None,
        "중복키수": None,
        "필수컬럼누락수": None,
        "주요수치결측수": None,
        "검증키": None
    }

    if df is None or df.empty:
        return result

    rule = rules.get(name, {})
    key_cols = rule.get("key", [])
    req_cols = rule.get("required", [])
    num_cols = rule.get("critical_numeric", [])

    if "연월" in df.columns and "연월_표시" in df.columns:
        result["연월파싱실패수"] = int(df["연월_표시"].isna().sum())

    missing_cols = [c for c in req_cols if c not in df.columns]
    result["필수컬럼누락수"] = len(missing_cols)
    result["검증키"] = " + ".join(key_cols) if key_cols else None

    if key_cols and all(c in df.columns for c in key_cols):
        result["중복키수"] = int(df.duplicated(key_cols).sum())

    exist_num_cols = [c for c in num_cols if c in df.columns]
    if exist_num_cols:
        result["주요수치결측수"] = int(df[exist_num_cols].isna().sum().sum())

    return result


# =========================================================
# 7. 앱 시작
# =========================================================
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
latest_month = months[-1] if months else None

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
    ]
)

# =========================================================
# 1. 종합 상황판
# =========================================================
if menu == "1. 종합 상황판":
    st.header("1. 종합 상황판")

    selected_dashboard_month = st.selectbox(
        "기준월 선택",
        months,
        index=len(months) - 1,
        key="dashboard_month"
    )

    curr = alert[alert["연월_표시"] == selected_dashboard_month].copy()

    prev_idx = months.index(selected_dashboard_month) - 1
    prev_month = months[prev_idx] if prev_idx >= 0 else None
    prev = alert[alert["연월_표시"] == prev_month].copy() if prev_month else pd.DataFrame()

    avg_curr = curr["최종위험점수"].mean() if "최종위험점수" in curr.columns else np.nan
    avg_prev = prev["최종위험점수"].mean() if not prev.empty and "최종위험점수" in prev.columns else np.nan
    delta_avg = avg_curr - avg_prev if pd.notna(avg_curr) and pd.notna(avg_prev) else None

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("기준월", selected_dashboard_month)
    c2.metric("분석 체인 수", len(curr))
    c3.metric(
        "우선관리대상 수",
        int((curr["상대적_우선관리대상"] == "Y").sum()) if "상대적_우선관리대상" in curr.columns else 0
    )
    c4.metric(
        "평균 최종위험점수",
        fmt_num(avg_curr, 2),
        delta=None if delta_avg is None else fmt_num(delta_avg, 2)
    )

    fig = px.bar(
        curr.sort_values("최종위험점수", ascending=False),
        x="체인구분",
        y="최종위험점수",
        color="최종경보등급",
        text_auto=".2f",
        title=f"{selected_dashboard_month} 체인별 최종위험점수"
    )
    fig.update_layout(height=420, xaxis_title="체인구분", yaxis_title="최종위험점수")
    st.plotly_chart(fig, use_container_width=True)

    show_cols = [
        c for c in [
            "체인구분", "최종위험점수", "최종경보등급",
            "보정사유", "대체조달가능성", "상대적_우선관리대상", "비고"
        ] if c in curr.columns
    ]
    st.dataframe(curr[show_cols], use_container_width=True)

    st.markdown("### 체인별 장기 비교")
    if compare_df is not None and not compare_df.empty:
        show_cols = [
            c for c in [
                "체인구분", "평균_최종위험점수", "최고 위험점수",
                "최고_위험_도달_시점_표시", "평균_FTA_활용비중",
                "평균_대체조달가능성점수", "우선관리대상 비중 (%)", "종합_시사점"
            ] if c in compare_df.columns
        ]
        st.dataframe(compare_df[show_cols], use_container_width=True)

# =========================================================
# 2. 체인별 심층 분석
# =========================================================
elif menu == "2. 체인별 심층 분석":
    st.header("2. 체인별 심층 분석")

    chain = st.selectbox("체인 선택", chains)
    p = panel[panel["체인구분"] == chain].sort_values("연월_sort").copy()
    a = alert[alert["체인구분"] == chain].sort_values("연월_sort").copy()
    latest = a.iloc[-1]

    prev_score = a.iloc[-2]["최종위험점수"] if len(a) >= 2 else np.nan
    latest_score = latest["최종위험점수"]
    delta_score = latest_score - prev_score if pd.notna(prev_score) else None

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("최신 최종위험점수", fmt_num(latest_score, 2), None if delta_score is None else fmt_num(delta_score, 2))
    k2.metric("최신 경보등급", latest["최종경보등급"])
    k3.metric("대체조달가능성", latest["대체조달가능성"])
    k4.metric("FTA 활용비중", fmt_num(latest["fta_ratio"], 2))

    fig1 = px.line(
        a, x="연월_표시", y="최종위험점수", markers=True,
        title=f"{chain} 최종위험점수 추이"
    )
    fig1.update_layout(height=320, xaxis_title="연월", yaxis_title="최종위험점수")
    st.plotly_chart(fig1, use_container_width=True)

    # 구조취약도: 정규화해서 표시
    structural_cols = [c for c in ["HHI", "상위1국의존도", "국가보정합계", "CV"] if c in p.columns]
    if structural_cols:
        norm_df = p[["연월_표시"] + structural_cols].copy()
        for c in structural_cols:
            norm_df[c] = normalize_0_100(norm_df[c])

        fig2 = px.line(
            norm_df,
            x="연월_표시",
            y=structural_cols,
            title=f"{chain} 구조 취약도 추이(정규화, 0~100)"
        )
        fig2.update_layout(height=320, xaxis_title="연월", yaxis_title="정규화 지표")
        st.plotly_chart(fig2, use_container_width=True)

    risk_cols = [c for c in ["가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"] if c in p.columns]
    if risk_cols:
        fig3 = px.line(
            p, x="연월_표시", y=risk_cols,
            title=f"{chain} 4대 리스크 추이"
        )
        fig3.update_layout(height=320, xaxis_title="연월", yaxis_title="리스크점수")
        st.plotly_chart(fig3, use_container_width=True)

    st.caption("※ 구조취약도 추이는 HHI·상위1국의존도·국가보정합계·CV의 단위 차이가 커서, 비교 가능하도록 0~100 정규화하여 표시했습니다.")

# =========================================================
# 3. 충격 원인 추적
# =========================================================
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
    risk_df = pd.DataFrame({
        "리스크유형": ["가격", "수급", "물류", "정책이벤트"],
        "점수": [
            a0.get("가격리스크점수"),
            a0.get("수급리스크점수"),
            a0.get("물류리스크점수"),
            a0.get("정책이벤트리스크점수"),
        ],
    }).sort_values("점수", ascending=False)

    fig = px.bar(
        risk_df, x="리스크유형", y="점수",
        color="리스크유형", text_auto=".2f"
    )
    fig.update_layout(height=350, xaxis_title="리스크유형", yaxis_title="점수")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 가격 내부요인 분해")
    if chain == "납산배터리군":
        detail = pd.DataFrame({
            "요인": ["환율정규화", "납가격정규화"],
            "정규화값": [p0.get("환율정규화"), p0.get("납가격정규화")],
            "5점값": [p0.get("환율정규화_5pt"), p0.get("납가격정규화_5pt")],
            "가중치(%)": [
                get_entropy_weight(entropy, chain, "환율정규화"),
                get_entropy_weight(entropy, chain, "납가격정규화"),
            ],
        })
    else:
        detail = pd.DataFrame({
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
        })

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

            base, rm, cf = build_country_risk_profile(country, risk_master, conflict, chain, month, code)

            st.caption("※ 국가별수입비중·보정점수는 선택한 연월 기준이며, 외부 국가리스크 점수(Composite/WGI/GPI/UCDP/ACLED)는 해당 연도 기준입니다.")

            if base:
                b1, b2, b3, b4 = st.columns(4)
                b1.metric("국가별수입비중", fmt_num(base.get("국가별수입비중"), 2))
                b2.metric("기본평가점수", fmt_num(base.get("기본평가점수"), 2))
                b3.metric("총보정점수", fmt_num(base.get("총보정점수"), 2))
                b4.metric("최종보정점수", fmt_num(base.get("최종보정점수"), 2))

                st.caption(build_country_explanation(base, rm))

            if rm is not None and not rm.empty:
                st.markdown("#### 국가 리스크 근거")
                show_rm = [
                    c for c in [
                        "Country", "Risk_Flag", "Composite_Risk_Score",
                        "WGI_Risk_Score_Used", "GPI_Risk_Score",
                        "UCDP_Risk_Score", "ACLED_Risk_Score"
                    ] if c in rm.columns
                ]
                st.dataframe(rm[show_rm], use_container_width=True)

            if cf is not None and not cf.empty:
                st.markdown("#### 최근 분쟁/갈등 근거")
                show_cf = [
                    c for c in ["국가명", "분쟁유형", "분쟁강도명", "불일치유형명", "시작일", "종료일"]
                    if c in cf.columns
                ]
                st.dataframe(cf[show_cf].head(10), use_container_width=True)

# =========================================================
# 4. 선행 신호 탐지
# =========================================================
elif menu == "4. 선행 신호 탐지":
    st.header("4. 선행 신호 탐지")

    chain = st.selectbox("체인 선택", chains, key="lead_chain")

    future_target = st.selectbox(
        "미래 반응지표 선택",
        ["최종위험점수", "가격리스크점수", "수급리스크점수", "물류리스크점수", "정책이벤트리스크점수"],
        index=0
    )

    p = panel[panel["체인구분"] == chain].sort_values("연월_sort").copy()

    lead = make_lead_table(p, chain, future_target)

    if lead.empty:
        st.warning("분석 가능한 선행신호가 부족합니다.")
    else:
        st.markdown("### 선행 신호 후보 요약")
        st.dataframe(lead, use_container_width=True)

        st.caption("해석 원칙: 본 메뉴는 lag 0을 제외하고, x(t)와 y(t+1~3)의 상관만 계산합니다. 즉 현재월 동행성이 아니라 향후 1~3개월 선행 신호만 탐지합니다.")

        meaningful = lead[lead["강도"].isin(["강함", "보통"])].copy()
        if meaningful.empty:
            st.info("현재 선택한 미래 반응지표에 대해 절댓값 기준 0.3 이상인 뚜렷한 선행 신호가 많지 않아, 참고용 수준으로 해석하는 것이 적절합니다.")

        xcol = st.selectbox("선행변수 선택", lead["선행변수"].tolist())
        corr_df = calc_forward_lag_corr(p, xcol, future_target, lags=(1, 2, 3))
        corr_df["강도"] = corr_df["corr"].apply(corr_strength_label)

        fig = px.bar(
            corr_df,
            x="lag_month",
            y="corr",
            color="강도",
            text_auto=".3f",
            title=f"{xcol}(t) → {future_target}(t+lag) 시차별 상관"
        )
        fig.update_layout(height=350, xaxis_title="선행 개월 수", yaxis_title="상관계수")
        st.plotly_chart(fig, use_container_width=True)

        best = corr_df.dropna(subset=["corr"])
        if not best.empty:
            best = best.loc[best["corr"].abs().idxmax()]
            st.write(build_lead_interpretation(xcol, future_target, int(best["lag_month"]), float(best["corr"])))

# =========================================================
# 5. 기업 대응 시뮬레이터
# =========================================================
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

    sim_df = pd.DataFrame({
        "지표": ["상위1국의존도", "HHI", "FTA 활용비중", "대체조달가능성점수", "최종위험점수"],
        "현재": [
            sim["현재 상위1국의존도"], sim["현재 HHI"], sim["현재 FTA 활용비중"],
            sim["현재 대체조달가능성점수"], sim["현재 최종위험점수"]
        ],
        "시뮬레이션": [
            sim["시뮬레이션 상위1국의존도"], sim["시뮬레이션 HHI"], sim["시뮬레이션 FTA 활용비중"],
            sim["시뮬레이션 대체조달가능성점수"], sim["시뮬레이션 최종위험점수"]
        ],
    })
    st.dataframe(sim_df, use_container_width=True)

    st.markdown("### 실행 해석")
    st.write("- 상위 공급국 집중 완화, FTA 활용 확대, 대체조달 역량 강화는 최종위험점수 완화에 유효한 방향입니다.")
    st.write("- 실무에서는 발주 분산, 신규 공급국 검토, 재고·비축 전략, 계약 구조 조정의 참고 시나리오로 활용할 수 있습니다.")

# =========================================================
# 6. 대체국 추천 시스템
# =========================================================
elif menu == "6. 대체국 추천 시스템":
    st.header("6. 대체국 추천 시스템")

    c1, c2 = st.columns(2)
    chain = c1.selectbox("체인", chains, key="alt_chain")
    month = c2.selectbox("연월", months, index=len(months) - 1, key="alt_month")

    keep_df, expand_df, reduce_df = recommend_country_groups(country, risk_master, chain, month)

    if keep_df.empty and expand_df.empty and reduce_df.empty:
        st.warning("추천 가능한 국가 데이터가 없습니다.")
    else:
        st.markdown("### 1) 유지 추천국")
        if not keep_df.empty:
            cols = [c for c in [
                "국가명", "국가코드", "국가수입금액", "국가별수입비중",
                "최종보정점수", "Composite_Risk_Score", "Risk_Flag",
                "FTA여부", "유지추천점수", "유지추천사유"
            ] if c in keep_df.columns]
            st.dataframe(keep_df[cols], use_container_width=True)
        else:
            st.info("유지 추천국 데이터가 없습니다.")

        st.markdown("### 2) 확대 후보국")
        if not expand_df.empty:
            cols = [c for c in [
                "국가명", "국가코드", "국가수입금액", "국가별수입비중",
                "최종보정점수", "Composite_Risk_Score", "Risk_Flag",
                "FTA여부", "확대추천점수", "확대추천사유"
            ] if c in expand_df.columns]
            st.dataframe(expand_df[cols], use_container_width=True)

            fig = px.bar(
                expand_df.sort_values("확대추천점수"),
                x="확대추천점수",
                y="국가명",
                orientation="h",
                text_auto=".2f",
                title=f"{chain} / {month} 확대 후보국 우선순위"
            )
            fig.update_layout(height=350, xaxis_title="확대추천점수", yaxis_title="국가명")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("확대 후보국 데이터가 없습니다.")

        st.markdown("### 3) 축소 검토국")
        if not reduce_df.empty:
            cols = [c for c in [
                "국가명", "국가코드", "국가수입금액", "국가별수입비중",
                "최종보정점수", "Composite_Risk_Score", "Risk_Flag",
                "FTA여부", "축소검토점수", "축소검토사유"
            ] if c in reduce_df.columns]
            st.dataframe(reduce_df[cols], use_container_width=True)
        else:
            st.info("축소 검토국 데이터가 없습니다.")

        st.markdown("### 추천 로직 요약")
        st.write(
            "- **유지 추천국**: 기존 거래 실적, 안정성, FTA 활용성을 종합 고려한 안정 조달국  \n"
            "- **확대 후보국**: 현재 비중이 과도하지 않으면서 집중도 완화 효과가 큰 국가  \n"
            "- **축소 검토국**: 현재 비중이 높고 내부·외부 위험 부담이 큰 국가"
        )

# =========================================================
# 7. 데이터 검증 / 방법론
# =========================================================
elif menu == "7. 데이터 검증 / 방법론":
    st.header("7. 데이터 검증 / 방법론")

    rules = get_validation_rules()
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
        rows.append(validate_sheet(name, df, rules))

    st.dataframe(pd.DataFrame(rows), use_container_width=True)

    st.markdown("### 검증 해석 유의사항")
    st.write("- `PANEL_MONTHLY`, `ALERT_RESULT`는 `연월 + 체인구분`이 유일키입니다.")
    st.write("- `COUNTRY_MONTHLY`는 국가·HS 단위 시트이므로 `연월 + 체인구분`만으로 중복을 보면 안 됩니다.")
    st.write("- `HS_MONTHLY_SUMMARY`도 HS 단위 시트이므로 시트별 고유키를 따로 적용합니다.")

    st.markdown("### 표준화 / 예외처리 반영사항")
    st.write("- `2021.01`, `2021-01`, `202101`을 `2021-01`로 통일")
    st.write("- 원본 특수값 `2021.1`은 `2021-10`으로 처리")
    st.write("- `GSCPI_NORM → GSCPI_Norm`, `이벤트보정 → TPU_Norm` 자동 통일")
    st.write("- float형 연월, Timestamp형 연월도 앱에서는 `YYYY-MM`으로 통일 표기")

    st.markdown("### 주요 산식")
    st.code(
        "CV = 국가별 월수입금액 population std / 평균\n"
        "경보점수기초 = (HHI/100 × 0.4) + (상위1국의존도 × 0.3) + (CV × 0.3)\n"
        "대체조달가능성_점수 = 0.35*수입국수_norm + 0.30*(100-상위1국의존도) + 0.20*(100-HHI_norm) + 0.15*fta_ratio\n"
        "선행 신호 탐지 = x(t)와 y(t+1~3)의 상관 비교 (lag 0 제외)"
    )

    st.markdown("### 해석 원칙")
    st.write("- 선행 신호 탐지는 인과 확정이 아니라 **향후 1~3개월 모니터링 우선순위**를 제시하는 보조 분석입니다.")
    st.write("- 대체국 추천은 안정성·FTA·기존 실적·집중도 완화 효과를 반영한 **1차 의사결정 지원**입니다.")
