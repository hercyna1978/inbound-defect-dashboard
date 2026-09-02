# -*- coding: utf-8 -*-
"""
입고 / 불량 데이터 대시보드
--------------------------------
주요 개선사항
1. 공장 / 불량타입 / 데이터 종류를 드롭다운(multiselect)에서 선택
   - 최초 실행 시 전체 항목 선택
   - 체크 해제로 원하는 항목만 남길 수 있음
2. 불량타입 필터는 '불량' 데이터에만 적용
   - 입고 데이터가 '기타'로 분류되어 입고 수량이 사라지는 기존 문제 방지
3. 월별 입고/불량 수량과 '월별 입고 대비 불량률'을 함께 표시
4. 누계 수량은 서로 다른 축을 겹쳐 그리지 않고,
   - 누계 입고 수량
   - 누계 불량 수량
   - 누계 불량률
   을 별도로 시각화
5. 불량타입 월별 그래프는 스택형이 아닌 그룹형 막대그래프
   - 테 / 렌즈를 메인으로 크게 표시
   - 전체 / 기타는 별도 '군소 타입' 탭에서 표시
6. SKU별 입고 대비 불량률 TOP 표 제공
   - 상품코드 기준으로 입고량과 불량량을 매칭
   - 최소 입고량 조건을 두어 입고 1~2개인 SKU가 100%로 상위에 뜨는 문제 완화
7. 원본 데이터 검증 정보 제공
   - 행 수, 수량 합계, 음수 수량, SKU 매칭 현황 등
"""

import io
import os
from typing import List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ------------------------------------------------------------------
# 기본 설정
# ------------------------------------------------------------------
st.set_page_config(
    page_title="입고 / 불량 데이터 대시보드",
    page_icon="📦",
    layout="wide",
)

DEFAULT_DATA_PATH = os.path.join(
    os.path.dirname(__file__), "data", "raw_data.xlsx"
)

FACTORY_COLORS = {
    "C5공장": "#1f77b4",
    "C2공장": "#2ca02c",
    "C2-S공장": "#ff7f0e",
    "미상": "#7f7f7f",
}

DEFECT_TYPE_COLORS = {
    "테": "#d62728",
    "렌즈": "#9467bd",
    "전체": "#8c564b",
    "기타": "#7f7f7f",
}

DEFECT_TYPES_ORDER = ["테", "렌즈", "전체", "기타"]
DATA_TYPES_ORDER = ["입고", "불량"]


# ------------------------------------------------------------------
# 분류 함수
# ------------------------------------------------------------------
def classify_factory(value) -> str:
    """공급처 상품명에서 공장을 분류한다."""
    if pd.isna(value) or str(value).strip() == "":
        return "미상"

    text = str(value).upper()

    # C2-S는 C2 문자열도 포함하므로 반드시 먼저 검사
    if "C2-S" in text:
        return "C2-S공장"
    if "C2" in text:
        return "C2공장"
    if "C5" in text:
        return "C5공장"

    return "미상"


def classify_defect_type(value) -> str:
    """
    전표제목 기준 불량타입 분류.

    중요:
    입고 행도 전표제목이 존재하므로,
    입고 행에는 이 분류를 필터 조건으로 적용하지 않는다.
    """
    if pd.isna(value) or str(value).strip() == "":
        return "기타"

    text = str(value)

    if "전체" in text:
        return "전체"
    if "테" in text:
        return "테"
    if "렌즈" in text:
        return "렌즈"

    return "기타"


# ------------------------------------------------------------------
# 데이터 로드
# ------------------------------------------------------------------
REQUIRED_COLUMNS = [
    "작업일",
    "수량",
    "공급처 상품명",
    "상품코드",
    "상품명",
    "전표제목",
]


@st.cache_data(show_spinner="엑셀 데이터를 불러오는 중입니다...")
def load_and_process(file_source) -> Tuple[pd.DataFrame, dict]:
    """불량/입고 시트를 읽고 공통 데이터프레임으로 만든다."""

    frames = []
    validation = {
        "dropped_invalid_dates": 0,
        "source_rows": {},
        "source_qty": {},
        "negative_qty_rows": {},
    }

    for sheet_name, data_kind in [("불량", "불량"), ("입고", "입고")]:
        df = pd.read_excel(file_source, sheet_name=sheet_name)

        validation["source_rows"][data_kind] = len(df)
        validation["source_qty"][data_kind] = pd.to_numeric(
            df["수량"], errors="coerce"
        ).fillna(0).sum()

        validation["negative_qty_rows"][data_kind] = int(
            (pd.to_numeric(df["수량"], errors="coerce").fillna(0) < 0).sum()
        )

        missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(
                f"'{sheet_name}' 시트에 필수 컬럼이 없습니다: {missing}"
            )

        df["출입구분"] = data_kind
        frames.append(df)

    df = pd.concat(frames, ignore_index=True, sort=False)

    before = len(df)
    df["작업일"] = pd.to_datetime(df["작업일"], errors="coerce")
    df = df.dropna(subset=["작업일"]).copy()
    validation["dropped_invalid_dates"] = before - len(df)

    df["년월"] = df["작업일"].dt.to_period("M").astype(str)

    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0)

    for col in ["상품명", "상품코드", "공급처", "공급처 상품명", "전표제목"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    df["공장"] = df["공급처 상품명"].apply(classify_factory)
    df["불량타입"] = df["전표제목"].apply(classify_defect_type)

    return df, validation


# ------------------------------------------------------------------
# 공통 함수
# ------------------------------------------------------------------
def safe_rate(numerator: float, denominator: float) -> float:
    """분모가 0이면 0을 반환."""
    return (numerator / denominator * 100.0) if denominator else 0.0


def filter_data(
    df: pd.DataFrame,
    selected_factories: List[str],
    selected_defect_types: List[str],
    selected_kinds: List[str],
    start_date,
    end_date,
    product_name_query: str,
    product_code_query: str,
) -> pd.DataFrame:
    """
    필터 적용.

    핵심:
    - 공장 / 데이터 종류는 모든 행에 적용
    - 불량타입은 불량 행에만 적용
    - 따라서 '테'만 선택해도 입고 데이터가 사라지지 않는다.
    """

    mask = (
        df["공장"].isin(selected_factories)
        & df["출입구분"].isin(selected_kinds)
        & (df["작업일"].dt.date >= start_date)
        & (df["작업일"].dt.date <= end_date)
    )

    result = df.loc[mask].copy()

    if not result.empty:
        defect_mask = result["출입구분"].eq("불량")
        result = result[
            (~defect_mask) | result["불량타입"].isin(selected_defect_types)
        ].copy()

    if product_name_query:
        result = result[
            result["상품명"].str.contains(
                product_name_query, case=False, na=False, regex=False
            )
        ].copy()

    if product_code_query:
        result = result[
            result["상품코드"].str.contains(
                product_code_query, case=False, na=False, regex=False
            )
        ].copy()

    return result


def monthly_summary(filtered: pd.DataFrame) -> pd.DataFrame:
    """월별 입고/불량 수량 및 불량률."""
    if filtered.empty:
        return pd.DataFrame(
            columns=["년월", "입고", "불량", "불량률"]
        )

    monthly = (
        filtered.groupby(["년월", "출입구분"])["수량"]
        .sum()
        .unstack(fill_value=0)
        .sort_index()
    )

    for col in DATA_TYPES_ORDER:
        if col not in monthly.columns:
            monthly[col] = 0

    monthly = monthly[DATA_TYPES_ORDER].copy()
    monthly["불량률"] = (
        monthly["불량"].div(monthly["입고"].replace(0, pd.NA)) * 100
    ).fillna(0)

    monthly.index.name = "년월"
    return monthly.reset_index()


def cumulative_summary(monthly: pd.DataFrame) -> pd.DataFrame:
    """월별 누계 및 누계 불량률."""
    result = monthly.copy()

    result["입고누계"] = result["입고"].cumsum()
    result["불량누계"] = result["불량"].cumsum()
    result["누계불량률"] = (
        result["불량누계"].div(result["입고누계"].replace(0, pd.NA)) * 100
    ).fillna(0)

    return result


def sku_defect_rate_table(
    filtered: pd.DataFrame,
    min_inbound_qty: int,
    top_n: int,
) -> pd.DataFrame:
    """
    상품코드(SKU) 기준 입고 대비 불량률.

    분모: 선택된 조건의 입고 수량
    분자: 선택된 조건의 불량 수량
    비율: 불량 / 입고 * 100
    """

    if filtered.empty:
        return pd.DataFrame()

    inbound = filtered[filtered["출입구분"] == "입고"].copy()
    defect = filtered[filtered["출입구분"] == "불량"].copy()

    inbound_group = (
        inbound.groupby("상품코드", as_index=False)
        .agg(
            입고수량=("수량", "sum"),
            상품명=("상품명", "first"),
            공장=("공장", lambda x: ", ".join(sorted(set(x)))),
        )
    )

    defect_group = (
        defect.groupby("상품코드", as_index=False)
        .agg(불량수량=("수량", "sum"))
    )

    if inbound_group.empty:
        return pd.DataFrame()

    result = inbound_group.merge(
        defect_group,
        on="상품코드",
        how="left",
    )

    result["불량수량"] = result["불량수량"].fillna(0)

    result = result[result["입고수량"] >= min_inbound_qty].copy()

    if result.empty:
        return result

    result["불량률"] = (
        result["불량수량"] / result["입고수량"] * 100
    )

    result = result.sort_values(
        ["불량률", "불량수량", "입고수량"],
        ascending=[False, False, False],
    ).head(top_n)

    result["순위"] = range(1, len(result) + 1)
    result["불량률"] = result["불량률"].round(2)

    return result[
        [
            "순위",
            "상품코드",
            "상품명",
            "공장",
            "입고수량",
            "불량수량",
            "불량률",
        ]
    ]


# ------------------------------------------------------------------
# 화면
# ------------------------------------------------------------------
st.title("📦 입고 / 불량 데이터 대시보드")
st.caption(
    "입고 대비 불량률을 중심으로 월별 추이, 불량타입, 공장, SKU 위험도를 확인합니다."
)


# ------------------------------------------------------------------
# 데이터 입력
# ------------------------------------------------------------------
data_source = None

if os.path.exists(DEFAULT_DATA_PATH):
    data_source = DEFAULT_DATA_PATH

uploaded = st.file_uploader(
    "raw_data.xlsx 업로드 (선택)",
    type=["xlsx"],
    help="기본 경로의 raw_data.xlsx가 있으면 자동으로 사용되며, 업로드한 파일이 있으면 업로드 파일을 사용합니다.",
)

if uploaded is not None:
    data_source = io.BytesIO(uploaded.getvalue())

if data_source is None:
    st.info(
        "data/raw_data.xlsx 파일을 넣거나 위에서 raw_data.xlsx를 업로드해주세요."
    )
    st.stop()

try:
    df, validation = load_and_process(data_source)
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()


# ------------------------------------------------------------------
# 데이터 검증 정보
# ------------------------------------------------------------------
with st.expander("🔎 원본 데이터 검증 결과", expanded=False):
    v1, v2, v3, v4 = st.columns(4)

    v1.metric(
        "불량 원본 행",
        f"{validation['source_rows'].get('불량', 0):,}",
    )
    v2.metric(
        "입고 원본 행",
        f"{validation['source_rows'].get('입고', 0):,}",
    )
    v3.metric(
        "불량 원본 수량",
        f"{validation['source_qty'].get('불량', 0):,.0f}",
    )
    v4.metric(
        "입고 원본 수량",
        f"{validation['source_qty'].get('입고', 0):,.0f}",
    )

    st.write(
        "음수 수량 행:",
        {
            "불량": validation["negative_qty_rows"].get("불량", 0),
            "입고": validation["negative_qty_rows"].get("입고", 0),
        },
    )

    if validation["dropped_invalid_dates"]:
        st.warning(
            f"작업일이 없거나 날짜로 변환되지 않아 "
            f"{validation['dropped_invalid_dates']:,}건이 집계에서 제외되었습니다."
        )
    else:
        st.success("작업일 누락/오류로 제외된 행은 없습니다.")

    defect_check = (
        df[df["출입구분"] == "불량"]
        .groupby("불량타입")["수량"]
        .agg(["count", "sum"])
        .reindex(DEFECT_TYPES_ORDER)
        .fillna(0)
    )

    st.write("불량타입별 원본 집계 검증")
    st.dataframe(
        defect_check.rename(
            columns={"count": "행수", "sum": "수량"}
        ).style.format({"행수": "{:,.0f}", "수량": "{:,.0f}"}),
        use_container_width=True,
    )

    st.caption(
        "※ 원본 불량 데이터는 전표제목 기준으로 테/렌즈/전체/기타로 분류합니다. "
        "음수 수량은 임의 삭제하지 않고 원본 순수량에 반영합니다."
    )


# ------------------------------------------------------------------
# 필터
# ------------------------------------------------------------------
st.subheader("🔎 필터")

filter1, filter2, filter3 = st.columns(3)

factory_options = sorted(df["공장"].dropna().unique().tolist())

with filter1:
    selected_factories = st.multiselect(
        "🏭 공장 선택",
        options=factory_options,
        default=factory_options,
        help="드롭다운에서 체크를 해제하여 원하는 공장만 선택할 수 있습니다.",
    )

with filter2:
    defect_type_options = [
        x for x in DEFECT_TYPES_ORDER if x in set(df["불량타입"])
    ]
    selected_defect_types = st.multiselect(
        "🏷️ 불량타입 선택",
        options=defect_type_options,
        default=defect_type_options,
        help="불량 데이터에만 적용됩니다. 입고 데이터는 불량타입 선택 때문에 제외되지 않습니다.",
    )

with filter3:
    selected_kinds = st.multiselect(
        "📂 데이터 종류 선택",
        options=DATA_TYPES_ORDER,
        default=DATA_TYPES_ORDER,
        help="입고 / 불량 중 원하는 데이터만 체크해 사용할 수 있습니다.",
    )

filter4, filter5, filter6 = st.columns([1.4, 1.4, 1])

with filter4:
    product_name_query = st.text_input(
        "상품명 검색",
        placeholder="예: BLACK, IREN",
    )

with filter5:
    product_code_query = st.text_input(
        "상품코드 검색",
        placeholder="예: 03595",
    )

with filter6:
    if df["작업일"].notna().any():
        min_date = df["작업일"].min().date()
        max_date = df["작업일"].max().date()
    else:
        st.error("유효한 작업일 데이터가 없습니다.")
        st.stop()

    date_range = st.date_input(
        "작업일 범위",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date = min_date
    end_date = max_date


# ------------------------------------------------------------------
# 선택값 검증
# ------------------------------------------------------------------
if not selected_factories:
    st.warning("공장을 하나 이상 선택해주세요.")
    st.stop()

if not selected_kinds:
    st.warning("데이터 종류를 하나 이상 선택해주세요.")
    st.stop()

if not selected_defect_types and "불량" in selected_kinds:
    st.warning("불량 데이터를 보려면 불량타입을 하나 이상 선택해주세요.")
    st.stop()


# ------------------------------------------------------------------
# 필터 적용
# ------------------------------------------------------------------
filtered = filter_data(
    df=df,
    selected_factories=selected_factories,
    selected_defect_types=selected_defect_types,
    selected_kinds=selected_kinds,
    start_date=start_date,
    end_date=end_date,
    product_name_query=product_name_query,
    product_code_query=product_code_query,
)

if filtered.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()


# ------------------------------------------------------------------
# KPI
# ------------------------------------------------------------------
total_in = filtered.loc[
    filtered["출입구분"] == "입고", "수량"
].sum()

total_defect = filtered.loc[
    filtered["출입구분"] == "불량", "수량"
].sum()

overall_rate = safe_rate(total_defect, total_in)

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

kpi1.metric("총 입고 수량", f"{total_in:,.0f}")
kpi2.metric("총 불량 수량", f"{total_defect:,.0f}")
kpi3.metric("입고 대비 불량률", f"{overall_rate:.2f}%")
kpi4.metric("필터링 행 수", f"{len(filtered):,}")
kpi5.metric("SKU 수", f"{filtered['상품코드'].nunique():,}")

st.divider()


# ------------------------------------------------------------------
# 월별 요약
# ------------------------------------------------------------------
monthly = monthly_summary(filtered)
cum = cumulative_summary(monthly)


# ------------------------------------------------------------------
# 1. 월별 입고/불량 + 불량률
# ------------------------------------------------------------------
st.subheader("📊 월별 입고 수량 / 불량 수량 / 입고 대비 불량률")

fig_month = go.Figure()

fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["입고"],
        name="입고 수량",
        text=monthly["입고"].map(lambda x: f"{x:,.0f}"),
        textposition="outside",
        marker_color="#1f77b4",
        opacity=0.75,
        hovertemplate="입고: %{y:,.0f}<extra></extra>",
    )
)

fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["불량"],
        name="불량 수량",
        text=monthly["불량"].map(lambda x: f"{x:,.0f}"),
        textposition="outside",
        marker_color="#d62728",
        opacity=0.85,
        hovertemplate="불량: %{y:,.0f}<extra></extra>",
    )
)

fig_month.add_trace(
    go.Scatter(
        x=monthly["년월"],
        y=monthly["불량률"],
        name="월별 불량률",
        mode="lines+markers+text",
        text=monthly["불량률"].map(lambda x: f"{x:.2f}%"),
        textposition="top center",
        line=dict(color="#111111", width=3),
        marker=dict(size=8),
        yaxis="y2",
        hovertemplate="불량률: %{y:.2f}%<extra></extra>",
    )
)

fig_month.update_layout(
    barmode="group",
    height=520,
    margin=dict(t=60, l=20, r=70, b=20),
    hovermode="x unified",
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
    ),
    xaxis=dict(title="년월"),
    yaxis=dict(
        title="수량",
        separatethousands=True,
    ),
    yaxis2=dict(
        title="불량률 (%)",
        overlaying="y",
        side="right",
        rangemode="tozero",
        ticksuffix="%",
    ),
)

st.plotly_chart(fig_month, use_container_width=True)

st.caption(
    "※ 수량은 막대, 불량률은 별도 라인으로 표시해 입고 대비 불량 수준을 한눈에 구분합니다."
)


# ------------------------------------------------------------------
# 2. 누계 추이
# ------------------------------------------------------------------
st.subheader("📈 월별 누계 추이")

cum_tab1, cum_tab2, cum_tab3 = st.tabs(
    ["누계 입고", "누계 불량", "누계 불량률"]
)

with cum_tab1:
    fig_cin = go.Figure()
    fig_cin.add_trace(
        go.Scatter(
            x=cum["년월"],
            y=cum["입고누계"],
            mode="lines+markers+text",
            text=cum["입고누계"].map(lambda x: f"{x:,.0f}"),
            textposition="top center",
            name="입고 누계",
            line=dict(color="#1f77b4", width=4),
        )
    )
    fig_cin.update_layout(
        height=420,
        xaxis_title="년월",
        yaxis_title="누계 입고 수량",
        hovermode="x unified",
        margin=dict(t=30, l=20, r=20, b=20),
    )
    st.plotly_chart(fig_cin, use_container_width=True)

with cum_tab2:
    fig_cdef = go.Figure()
    fig_cdef.add_trace(
        go.Scatter(
            x=cum["년월"],
            y=cum["불량누계"],
            mode="lines+markers+text",
            text=cum["불량누계"].map(lambda x: f"{x:,.0f}"),
            textposition="top center",
            name="불량 누계",
            line=dict(color="#d62728", width=4),
        )
    )
    fig_cdef.update_layout(
        height=420,
        xaxis_title="년월",
        yaxis_title="누계 불량 수량",
        hovermode="x unified",
        margin=dict(t=30, l=20, r=20, b=20),
    )
    st.plotly_chart(fig_cdef, use_container_width=True)

with cum_tab3:
    fig_crate = go.Figure()
    fig_crate.add_trace(
        go.Scatter(
            x=cum["년월"],
            y=cum["누계불량률"],
            mode="lines+markers+text",
            text=cum["누계불량률"].map(lambda x: f"{x:.2f}%"),
            textposition="top center",
            name="누계 불량률",
            line=dict(color="#111111", width=4),
            fill="tozeroy",
        )
    )
    fig_crate.update_layout(
        height=420,
        xaxis_title="년월",
        yaxis_title="누계 불량률 (%)",
        hovermode="x unified",
        margin=dict(t=30, l=20, r=20, b=20),
        yaxis=dict(ticksuffix="%"),
    )
    st.plotly_chart(fig_crate, use_container_width=True)

st.caption(
    "※ 기존처럼 좌우 축에 입고/불량 누계 수량을 겹쳐 그리지 않고 각각 분리했습니다."
)


# ------------------------------------------------------------------
# 3. 불량타입 월별 그래프
# ------------------------------------------------------------------
st.divider()
st.subheader("🏷️ 불량타입별 월별 수량")

defect_only = filtered[filtered["출입구분"] == "불량"].copy()

if defect_only.empty:
    st.info("현재 조건에서는 불량 데이터가 없습니다.")
else:
    type_monthly = (
        defect_only.groupby(["년월", "불량타입"])["수량"]
        .sum()
        .unstack(fill_value=0)
        .sort_index()
    )

    for col in DEFECT_TYPES_ORDER:
        if col not in type_monthly.columns:
            type_monthly[col] = 0

    type_monthly = type_monthly[DEFECT_TYPES_ORDER]

    tab_main, tab_minor, tab_table = st.tabs(
        ["주요 타입 · 테 / 렌즈", "군소 타입 · 전체 / 기타", "수량 검증표"]
    )

    def make_defect_group_chart(
        data: pd.DataFrame,
        types: List[str],
        title: str,
    ):
        fig = go.Figure()

        for defect_type in types:
            if defect_type not in data.columns:
                continue

            fig.add_trace(
                go.Bar(
                    x=data.index,
                    y=data[defect_type],
                    name=defect_type,
                    marker_color=DEFECT_TYPE_COLORS[defect_type],
                    text=data[defect_type].map(
                        lambda x: f"{x:,.0f}" if x != 0 else ""
                    ),
                    textposition="outside",
                    hovertemplate=(
                        f"{defect_type}: "
                        "%{y:,.0f}<extra></extra>"
                    ),
                )
            )

        fig.update_layout(
            title=title,
            barmode="group",
            height=500,
            xaxis_title="년월",
            yaxis_title="불량 수량",
            hovermode="x unified",
            margin=dict(t=60, l=20, r=20, b=20),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
            ),
        )

        return fig

    with tab_main:
        main_types = [
            x for x in ["테", "렌즈"]
            if x in type_monthly.columns
            and (type_monthly[x].abs().sum() != 0)
        ]

        if main_types:
            st.plotly_chart(
                make_defect_group_chart(
                    type_monthly,
                    main_types,
                    "테 / 렌즈 월별 불량 수량",
                ),
                use_container_width=True,
            )
        else:
            st.info("테/렌즈 데이터가 없습니다.")

    with tab_minor:
        minor_types = [
            x for x in ["전체", "기타"]
            if x in type_monthly.columns
            and (type_monthly[x].abs().sum() != 0)
        ]

        if minor_types:
            st.plotly_chart(
                make_defect_group_chart(
                    type_monthly,
                    minor_types,
                    "전체 / 기타 월별 불량 수량",
                ),
                use_container_width=True,
            )
        else:
            st.info(
                "현재 선택 조건에서는 전체/기타 불량 수량이 없습니다."
            )

    with tab_table:
        check_table = type_monthly.copy()
        check_table["불량 합계"] = check_table.sum(axis=1)

        st.dataframe(
            check_table.style.format("{:,.0f}"),
            use_container_width=True,
        )

        st.caption(
            "※ 스택형 그래프를 제거했습니다. 각 타입의 막대가 같은 기준선에서 나란히 표시됩니다."
        )


# ------------------------------------------------------------------
# 4. 공장별 현황
# ------------------------------------------------------------------
st.divider()
st.subheader("🏭 공장별 입고 / 불량 / 불량률")

factory_summary = (
    filtered.groupby(["공장", "출입구분"])["수량"]
    .sum()
    .unstack(fill_value=0)
)

for col in DATA_TYPES_ORDER:
    if col not in factory_summary.columns:
        factory_summary[col] = 0

factory_summary = factory_summary[DATA_TYPES_ORDER].copy()
factory_summary["불량률"] = (
    factory_summary["불량"]
    .div(factory_summary["입고"].replace(0, pd.NA))
    * 100
).fillna(0)

fig_factory = go.Figure()

fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["입고"],
        name="입고",
        marker_color=[
            FACTORY_COLORS.get(x, "#7f7f7f")
            for x in factory_summary.index
        ],
    )
)

fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["불량"],
        name="불량",
        marker_color="#d62728",
    )
)

fig_factory.add_trace(
    go.Scatter(
        x=factory_summary.index,
        y=factory_summary["불량률"],
        name="불량률",
        mode="lines+markers+text",
        text=factory_summary["불량률"].map(lambda x: f"{x:.2f}%"),
        textposition="top center",
        yaxis="y2",
        line=dict(color="#111111", width=3),
    )
)

fig_factory.update_layout(
    barmode="group",
    height=480,
    xaxis_title="공장",
    yaxis=dict(title="수량"),
    yaxis2=dict(
        title="불량률 (%)",
        overlaying="y",
        side="right",
        rangemode="tozero",
        ticksuffix="%",
    ),
    hovermode="x unified",
    margin=dict(t=40, l=20, r=70, b=20),
)

st.plotly_chart(fig_factory, use_container_width=True)

st.dataframe(
    factory_summary.style.format(
        {
            "입고": "{:,.0f}",
            "불량": "{:,.0f}",
            "불량률": "{:.2f}%",
        }
    ),
    use_container_width=True,
)


# ------------------------------------------------------------------
# 5. SKU별 입고 대비 불량률
# ------------------------------------------------------------------
st.divider()
st.subheader("🚨 입고 대비 불량률이 높은 SKU")

sku_col1, sku_col2 = st.columns([1, 1])

with sku_col1:
    min_inbound_qty = st.number_input(
        "최소 입고 수량",
        min_value=1,
        value=100,
        step=10,
        help=(
            "입고 수량이 너무 작은 SKU가 불량 1~2개만으로 "
            "불량률 100%에 가까워지는 것을 방지합니다."
        ),
    )

with sku_col2:
    top_n = st.number_input(
        "표시 SKU 수",
        min_value=5,
        max_value=100,
        value=20,
        step=5,
    )

sku_table = sku_defect_rate_table(
    filtered=filtered,
    min_inbound_qty=int(min_inbound_qty),
    top_n=int(top_n),
)

if sku_table.empty:
    st.info(
        "조건을 만족하는 SKU가 없습니다. 최소 입고 수량을 낮춰보세요."
    )
else:
    st.dataframe(
        sku_table.style.format(
            {
                "입고수량": "{:,.0f}",
                "불량수량": "{:,.0f}",
                "불량률": "{:.2f}%",
            }
        ),
        use_container_width=True,
        height=520,
    )

    st.caption(
        "※ SKU 불량률 = 선택된 조건의 불량 수량 ÷ 선택된 조건의 입고 수량 × 100. "
        "상품코드가 같은 입고/불량 데이터를 연결하여 계산합니다."
    )


# ------------------------------------------------------------------
# 6. 불량타입별 총합
# ------------------------------------------------------------------
st.divider()
st.subheader("📋 불량타입별 총 수량 / 비중")

type_total = (
    defect_only.groupby("불량타입")["수량"]
    .sum()
    .reindex(DEFECT_TYPES_ORDER)
    .fillna(0)
    .to_frame("불량수량")
)

type_total["비중"] = (
    type_total["불량수량"] /
    type_total["불량수량"].sum() * 100
).fillna(0)

type_total["누계"] = type_total["불량수량"].cumsum()

st.dataframe(
    type_total.style.format(
        {
            "불량수량": "{:,.0f}",
            "비중": "{:.2f}%",
            "누계": "{:,.0f}",
        }
    ),
    use_container_width=True,
)


# ------------------------------------------------------------------
# 7. 상세 데이터
# ------------------------------------------------------------------
st.divider()
st.subheader("📋 상세 데이터")

display_cols = [
    "작업일",
    "년월",
    "출입구분",
    "공장",
    "불량타입",
    "공급처",
    "공급처 상품명",
    "상품코드",
    "상품명",
    "옵션",
    "수량",
    "전표제목",
    "전표번호",
]

display_cols = [
    c for c in display_cols if c in filtered.columns
]

detail = filtered[display_cols].sort_values(
    "작업일",
    ascending=False,
)

st.dataframe(
    detail,
    use_container_width=True,
    height=450,
)

csv_bytes = detail.to_csv(
    index=False
).encode("utf-8-sig")

st.download_button(
    label="⬇️ 현재 필터 데이터 CSV 다운로드",
    data=csv_bytes,
    file_name="filtered_data.csv",
    mime="text/csv",
)
