# -*- coding: utf-8 -*-
"""
입고 / 불량 데이터 대시보드
--------------------------------
- '불량', '입고' 두 시트를 읽어 하나의 데이터프레임으로 통합
- 작업일 기준으로 월(년월)을 추출
- 공급처 상품명을 기준으로 공장을 분류 (C2-S / C2 / C5 / 미상)
- 전표제목을 기준으로 불량타입을 분류 (테 / 렌즈 / 전체 / 기타)
- 상단 필터(공장, 불량타입, 상품명, 상품코드, 데이터종류, 작업일 범위)로 데이터를 좁혀가며
  월별 누계 수량 추이, 불량타입별 월별 스택 그래프, 공장별/불량타입별 고정 색상 그래프를
  확인할 수 있는 Streamlit 대시보드
"""

import io
import os

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# ------------------------------------------------------------------
# 기본 설정
# ------------------------------------------------------------------
st.set_page_config(
    page_title="입고/불량 데이터 대시보드",
    page_icon="📦",
    layout="wide",
)

DEFAULT_DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "raw_data.xlsx")

# 공장별 고정 색상 (모든 그래프에서 동일하게 사용)
FACTORY_COLORS = {
    "C5공장": "#1f77b4",
    "C2공장": "#2ca02c",
    "C2-S공장": "#ff7f0e",
    "미상": "#7f7f7f",
}

# 불량타입별 고정 색상 (모든 그래프에서 동일하게 사용)
DEFECT_TYPE_COLORS = {
    "테": "#d62728",
    "렌즈": "#9467bd",
    "전체": "#8c564b",
    "기타": "#7f7f7f",
}


# ------------------------------------------------------------------
# 분류 규칙
# ------------------------------------------------------------------
def classify_factory(supplier_product_name: str) -> str:
    """공급처 상품명 값을 보고 공장을 분류한다.

    - 값이 비어있으면 '미상'
    - 대소문자 구분 없이 'C2-S' 가 포함되면 'C2-S공장'
    - 대소문자 구분 없이 'C2' 가 포함되면 'C2공장'
    - 대소문자 구분 없이 'C5' 가 포함되면 'C5공장'
    - 위에 해당하지 않으면 '미상'
    """
    if pd.isna(supplier_product_name) or str(supplier_product_name).strip() == "":
        return "미상"

    s = str(supplier_product_name).upper()

    # C2-S 는 C2 보다 먼저 체크해야 함 (C2-S 안에 C2 문자열이 포함되므로)
    if "C2-S" in s:
        return "C2-S공장"
    if "C2" in s:
        return "C2공장"
    if "C5" in s:
        return "C5공장"
    return "미상"


def classify_defect_type(title: str) -> str:
    """전표제목을 보고 테 / 렌즈 / 전체 / 기타 로 분류한다."""
    if pd.isna(title) or str(title).strip() == "":
        return "기타"

    t = str(title)
    if "전체" in t:
        return "전체"
    if "테" in t:
        return "테"
    if "렌즈" in t:
        return "렌즈"
    return "기타"


# ------------------------------------------------------------------
# 데이터 로드 & 전처리
# ------------------------------------------------------------------
@st.cache_data(show_spinner="데이터를 불러오는 중입니다...")
def load_and_process(file) -> pd.DataFrame:
    frames = []
    for sheet_name, kind in [("불량", "불량"), ("입고", "입고")]:
        df = pd.read_excel(file, sheet_name=sheet_name)
        df["출입구분"] = kind
        frames.append(df)

    df = pd.concat(frames, ignore_index=True, sort=False)

    # 작업일 파싱
    df["작업일"] = pd.to_datetime(df["작업일"], errors="coerce")
    before = len(df)
    df = df.dropna(subset=["작업일"])
    dropped = before - len(df)

    # 월(년월) 추출
    df["년월"] = df["작업일"].dt.to_period("M").astype(str)

    # 공장 분류
    df["공장"] = df["공급처 상품명"].apply(classify_factory)

    # 불량타입(테/렌즈/전체/기타) 분류 - 전표제목 기준
    df["불량타입"] = df["전표제목"].apply(classify_defect_type)

    # 수량 결측치 처리
    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0)

    # 검색 편의를 위한 컬럼 보정
    for col in ["상품명", "상품코드", "공급처"]:
        if col in df.columns:
            df[col] = df[col].astype(str).replace("nan", "")

    df.attrs["dropped_rows"] = dropped
    return df


# ------------------------------------------------------------------
# 데이터 입력 (기본 경로에 파일이 있으면 자동 로드, 없으면 업로드 요청)
# ------------------------------------------------------------------
st.title("📦 입고 / 불량 데이터 대시보드")
st.caption("작업일 기준 월별 입고·불량 수량 추이와 공장별/유형별 세부 현황을 확인합니다.")

data_source = None
if os.path.exists(DEFAULT_DATA_PATH):
    data_source = DEFAULT_DATA_PATH
else:
    uploaded = st.file_uploader(
        "raw_data.xlsx 파일을 업로드하세요 ('불량', '입고' 시트가 포함된 파일)",
        type=["xlsx"],
    )
    if uploaded is not None:
        data_source = io.BytesIO(uploaded.read())

if data_source is None:
    st.info("좌측 또는 상단에서 데이터 파일을 업로드하면 대시보드가 표시됩니다.")
    st.stop()

df = load_and_process(data_source)

if df.attrs.get("dropped_rows"):
    st.warning(f"작업일 값이 비어있거나 형식이 올바르지 않은 {df.attrs['dropped_rows']}건은 집계에서 제외되었습니다.")

# ------------------------------------------------------------------
# 상단 필터 영역
# ------------------------------------------------------------------
st.subheader("🔎 필터")

filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([1.2, 1.2, 1, 1])

with filter_col1:
    factory_options = sorted(df["공장"].unique().tolist())
    selected_factories = st.multiselect(
        "공장별 선택",
        options=factory_options,
        default=factory_options,
    )

with filter_col2:
    defect_type_options = ["테", "렌즈", "전체", "기타"]
    defect_type_options = [d for d in defect_type_options if d in df["불량타입"].unique()]
    selected_defect_types = st.multiselect(
        "불량타입(테/렌즈/전체/기타)",
        options=defect_type_options,
        default=defect_type_options,
    )

with filter_col3:
    product_name_query = st.text_input("상품명 검색", placeholder="예: BLACK, IREN ...")

with filter_col4:
    product_code_query = st.text_input("상품코드 검색", placeholder="예: 03595")

# 출입구분(입고/불량) 필터 + 날짜 범위 필터
filter_col5, filter_col6 = st.columns([1, 2])

with filter_col5:
    kind_options = ["입고", "불량"]
    selected_kinds = st.multiselect("데이터 종류", options=kind_options, default=kind_options)

with filter_col6:
    min_date = df["작업일"].min().date()
    max_date = df["작업일"].max().date()
    date_range = st.date_input(
        "작업일 범위",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

# 날짜를 하나만 선택한 경우(범위 선택이 끝나지 않은 경우) 방지
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

# ------------------------------------------------------------------
# 필터 적용
# ------------------------------------------------------------------
filtered = df[
    df["공장"].isin(selected_factories)
    & df["불량타입"].isin(selected_defect_types)
    & df["출입구분"].isin(selected_kinds)
    & (df["작업일"].dt.date >= start_date)
    & (df["작업일"].dt.date <= end_date)
]

if product_name_query:
    filtered = filtered[
        filtered["상품명"].str.contains(product_name_query, case=False, na=False)
    ]

if product_code_query:
    filtered = filtered[
        filtered["상품코드"].astype(str).str.contains(product_code_query, case=False, na=False)
    ]

if filtered.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다. 필터를 조정해주세요.")
    st.stop()

# ------------------------------------------------------------------
# KPI 요약
# ------------------------------------------------------------------
total_in = filtered.loc[filtered["출입구분"] == "입고", "수량"].sum()
total_defect = filtered.loc[filtered["출입구분"] == "불량", "수량"].sum()
defect_rate = (total_defect / total_in * 100) if total_in else 0

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("총 입고 수량", f"{total_in:,.0f}")
kpi2.metric("총 불량 수량", f"{total_defect:,.0f}")
kpi3.metric("불량률", f"{defect_rate:.2f}%")
kpi4.metric("필터링된 행 수", f"{len(filtered):,}")

st.divider()

# ------------------------------------------------------------------
# 월별 누계 수량 추이 그래프
# ------------------------------------------------------------------
st.subheader("📈 월별 누계 수량 추이 (입고 vs 불량)")

monthly = (
    filtered.groupby(["년월", "출입구분"])["수량"].sum().unstack(fill_value=0).sort_index()
)
for col in ["입고", "불량"]:
    if col not in monthly.columns:
        monthly[col] = 0
monthly_cum = monthly[["입고", "불량"]].cumsum()

fig = make_subplots(specs=[[{"secondary_y": True}]])
fig.add_trace(
    go.Scatter(
        x=monthly_cum.index,
        y=monthly_cum["입고"],
        name="입고 누계 수량",
        mode="lines+markers",
        line=dict(color="#1f77b4", width=3),
    ),
    secondary_y=False,
)
fig.add_trace(
    go.Scatter(
        x=monthly_cum.index,
        y=monthly_cum["불량"],
        name="불량 누계 수량",
        mode="lines+markers",
        line=dict(color="#d62728", width=3),
    ),
    secondary_y=True,
)
fig.update_layout(
    hovermode="x unified",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    margin=dict(t=40, l=10, r=10, b=10),
)
fig.update_xaxes(title_text="년월")
fig.update_yaxes(title_text="입고 누계 수량", secondary_y=False)
fig.update_yaxes(title_text="불량 누계 수량", secondary_y=True)

st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------------
# 월별 단순 수량(비누계) 참고 그래프
# ------------------------------------------------------------------
with st.expander("📊 월별 수량(누계 아님) 그래프 보기"):
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=monthly.index, y=monthly["입고"], name="입고 수량", marker_color="#1f77b4"))
    fig2.add_trace(go.Bar(x=monthly.index, y=monthly["불량"], name="불량 수량", marker_color="#d62728"))
    fig2.update_layout(barmode="group", xaxis_title="년월", yaxis_title="수량", margin=dict(t=20, l=10, r=10, b=10))
    st.plotly_chart(fig2, use_container_width=True)

# ------------------------------------------------------------------
# 불량타입별 월별 스택 그래프 (불량 데이터 기준)
# ------------------------------------------------------------------
st.subheader("📚 불량타입별 월별 수량 (스택 그래프)")

defect_only = filtered[filtered["출입구분"] == "불량"]
if defect_only.empty:
    st.info("현재 필터 조건에서는 '불량' 데이터가 없어 스택 그래프를 표시할 수 없습니다.")
else:
    type_monthly = (
        defect_only.groupby(["년월", "불량타입"])["수량"].sum().unstack(fill_value=0).sort_index()
    )
    fig3 = go.Figure()
    for t in ["테", "렌즈", "전체", "기타"]:
        if t in type_monthly.columns:
            fig3.add_trace(
                go.Bar(
                    x=type_monthly.index,
                    y=type_monthly[t],
                    name=t,
                    marker_color=DEFECT_TYPE_COLORS.get(t),
                )
            )
    fig3.update_layout(
        barmode="stack",
        xaxis_title="년월",
        yaxis_title="불량 수량",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=40, l=10, r=10, b=10),
    )
    st.plotly_chart(fig3, use_container_width=True)

st.divider()

# ------------------------------------------------------------------
# 공장별 / 불량타입별 세부 현황
# ------------------------------------------------------------------
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("🏭 공장별 수량")
    factory_summary = (
        filtered.groupby(["공장", "출입구분"])["수량"].sum().unstack(fill_value=0)
    )
    for col in ["입고", "불량"]:
        if col not in factory_summary.columns:
            factory_summary[col] = 0
    factory_summary = factory_summary[["입고", "불량"]]

    fig4 = go.Figure()
    for kind, pattern in [("입고", ""), ("불량", "/")]:
        fig4.add_trace(
            go.Bar(
                x=factory_summary.index,
                y=factory_summary[kind],
                name=kind,
                marker=dict(
                    color=[FACTORY_COLORS.get(f, "#7f7f7f") for f in factory_summary.index],
                    pattern_shape=pattern,
                ),
            )
        )
    fig4.update_layout(
        barmode="group",
        xaxis_title="공장",
        yaxis_title="수량",
        margin=dict(t=10, l=10, r=10, b=10),
    )
    st.plotly_chart(fig4, use_container_width=True)
    st.dataframe(factory_summary, use_container_width=True)

with col_b:
    st.subheader("🏷️ 불량타입별 수량 (테/렌즈/전체/기타)")
    type_summary = (
        filtered.groupby(["불량타입", "출입구분"])["수량"].sum().unstack(fill_value=0)
    )
    for col in ["입고", "불량"]:
        if col not in type_summary.columns:
            type_summary[col] = 0
    type_summary = type_summary[["입고", "불량"]]

    fig5 = go.Figure()
    for kind, pattern in [("입고", ""), ("불량", "/")]:
        fig5.add_trace(
            go.Bar(
                x=type_summary.index,
                y=type_summary[kind],
                name=kind,
                marker=dict(
                    color=[DEFECT_TYPE_COLORS.get(t, "#7f7f7f") for t in type_summary.index],
                    pattern_shape=pattern,
                ),
            )
        )
    fig5.update_layout(
        barmode="group",
        xaxis_title="불량타입",
        yaxis_title="수량",
        margin=dict(t=10, l=10, r=10, b=10),
    )
    st.plotly_chart(fig5, use_container_width=True)
    st.dataframe(type_summary, use_container_width=True)

st.divider()

# ------------------------------------------------------------------
# 상세 데이터 테이블 & 다운로드
# ------------------------------------------------------------------
st.subheader("📋 상세 데이터")

display_cols = [
    "작업일", "년월", "출입구분", "공장", "불량타입",
    "공급처", "공급처 상품명", "상품코드", "상품명", "옵션", "수량", "전표제목", "전표번호",
]
display_cols = [c for c in display_cols if c in filtered.columns]

st.dataframe(
    filtered[display_cols].sort_values("작업일", ascending=False),
    use_container_width=True,
    height=400,
)

csv_bytes = filtered[display_cols].to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="⬇️ 필터링된 데이터 CSV 다운로드",
    data=csv_bytes,
    file_name="filtered_data.csv",
    mime="text/csv",
)
