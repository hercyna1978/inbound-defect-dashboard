# -*- coding: utf-8 -*-
"""
입고 / 불량 / 기초재고 대시보드
================================
핵심 기준
1) 2026-01-01 기초재고를 출발점으로 사용한다.
2) 재고 흐름 = 기초재고 + 누적 입고 - 누적 불량
3) 불량률의 기본 분모는 '입고만'이 아니라 '당월 가용재고'이다.
   - 당월 가용재고 = 월초 이론재고 + 당월 입고
   - 누계 불량률 = 누적 불량 / (기초재고 + 누적 입고)
4) SKU별 위험도도 기초재고 + 누적 입고를 분모로 계산한다.
5) 불량 > 입고인 현상은 기초재고가 있으면 정상적으로 발생할 수 있으므로
   '입고량보다 불량이 많다 = 오류'로 판정하지 않는다.
6) 다만 기초재고 + 누적 입고 - 누적 불량이 음수가 되는 SKU는
   '재고흐름 점검' 대상으로 별도 표시한다.
   (판매/출고/이동 등 이 파일에 없는 재고 감소 거래가 있으면 음수가
    생길 수 있으므로 자동으로 원본 오류라고 단정하지 않는다.)
7) 필터는 Excel 필터처럼 multiselect 체크 목록으로 제공하며
   최초에는 전체 선택 상태이다.
8) Plotly 색상/템플릿은 Streamlit의 현재 테마를 감지하여
   라이트/다크 모드에 맞게 자동 변경한다.
"""

import io
import os
from typing import List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="입고 / 불량 / 기초재고 대시보드",
    page_icon="📦",
    layout="wide",
)

DEFAULT_DATA_PATH = os.path.join(
    os.path.dirname(__file__), "data", "raw_data.xlsx"
)

FACTORY_COLORS_LIGHT = {
    "C5공장": "#1565C0",
    "C2공장": "#2E7D32",
    "C2-S공장": "#EF6C00",
    "미상": "#757575",
}
DEFECT_TYPE_COLORS_LIGHT = {
    "테": "#D32F2F",
    "렌즈": "#7B1FA2",
    "전체": "#6D4C41",
    "기타": "#616161",
}
FACTORY_COLORS_DARK = {
    "C5공장": "#64B5F6",
    "C2공장": "#81C784",
    "C2-S공장": "#FFB74D",
    "미상": "#BDBDBD",
}
DEFECT_TYPE_COLORS_DARK = {
    "테": "#EF5350",
    "렌즈": "#CE93D8",
    "전체": "#BCAAA4",
    "기타": "#B0BEC5",
}

DEFECT_TYPES_ORDER = ["테", "렌즈", "전체", "기타"]
DATA_TYPES_ORDER = ["입고", "불량"]
REQUIRED_TRANSACTION_COLUMNS = [
    "작업일", "수량", "공급처 상품명", "상품코드", "상품명", "전표제목"
]
REQUIRED_BASE_COLUMNS = ["기준일", "상품코드", "상품명", "현재고수량"]


def get_theme_type() -> str:
    """Streamlit 현재 테마를 최대한 안전하게 감지한다."""
    try:
        theme = getattr(st.context, "theme", None)
        theme_type = getattr(theme, "type", None)
        if theme_type in ("dark", "light"):
            return theme_type
    except Exception:
        pass

    try:
        base = st.get_option("theme.base")
        if base in ("dark", "light"):
            return base
    except Exception:
        pass

    # 테마 정보를 읽지 못하면 기본값은 light.
    return "light"


def chart_colors():
    if get_theme_type() == "dark":
        return FACTORY_COLORS_DARK, DEFECT_TYPE_COLORS_DARK
    return FACTORY_COLORS_LIGHT, DEFECT_TYPE_COLORS_LIGHT


def apply_theme(fig: go.Figure, height: int = 480):
    """Plotly를 Streamlit 라이트/다크 모드에 맞춘다."""
    dark = get_theme_type() == "dark"
    template = "plotly_dark" if dark else "plotly_white"
    grid_color = "rgba(255,255,255,0.15)" if dark else "rgba(0,0,0,0.10)"
    zero_color = "rgba(255,255,255,0.35)" if dark else "rgba(0,0,0,0.25)"

    fig.update_layout(
        template=template,
        height=height,
        margin=dict(t=55, l=20, r=75, b=35),
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
        font=dict(size=13),
        xaxis=dict(
            showgrid=False,
            zeroline=False,
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor=grid_color,
            zeroline=True,
            zerolinecolor=zero_color,
        ),
    )
    return fig


def classify_factory(value) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return "미상"

    text = str(value).upper()
    if "C2-S" in text:
        return "C2-S공장"
    if "C2" in text:
        return "C2공장"
    if "C5" in text:
        return "C5공장"
    return "미상"


def classify_defect_type(value) -> str:
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


@st.cache_data(show_spinner="엑셀 데이터를 불러오고 검증하는 중입니다...")
def load_and_process(file_source) -> Tuple[pd.DataFrame, pd.DataFrame, dict]:
    frames = []
    validation = {
        "source_rows": {},
        "source_qty": {},
        "negative_qty_rows": {},
        "invalid_dates": {},
        "base_rows": 0,
        "base_qty": 0,
        "base_date_values": [],
    }

    for sheet_name, data_kind in [("불량", "불량"), ("입고", "입고")]:
        df = pd.read_excel(file_source, sheet_name=sheet_name)

        missing = [c for c in REQUIRED_TRANSACTION_COLUMNS if c not in df.columns]
        if missing:
            raise ValueError(
                f"'{sheet_name}' 시트에 필수 컬럼이 없습니다: {missing}"
            )

        qty = pd.to_numeric(df["수량"], errors="coerce").fillna(0)
        work_date = pd.to_datetime(df["작업일"], errors="coerce")

        validation["source_rows"][data_kind] = len(df)
        validation["source_qty"][data_kind] = float(qty.sum())
        validation["negative_qty_rows"][data_kind] = int((qty < 0).sum())
        validation["invalid_dates"][data_kind] = int(work_date.isna().sum())

        df["출입구분"] = data_kind
        df["수량"] = qty
        df["작업일"] = work_date
        frames.append(df)

    try:
        base = pd.read_excel(file_source, sheet_name="기초재고")
    except Exception as e:
        raise ValueError(
            "이번 버전은 반드시 '기초재고' 시트가 필요합니다. "
            f"기초재고 시트를 읽을 수 없습니다: {e}"
        )

    missing_base = [c for c in REQUIRED_BASE_COLUMNS if c not in base.columns]
    if missing_base:
        raise ValueError(
            f"'기초재고' 시트에 필수 컬럼이 없습니다: {missing_base}"
        )

    base["기준일"] = pd.to_datetime(base["기준일"], errors="coerce")
    base["상품코드"] = base["상품코드"].fillna("").astype(str).str.strip()
    base["상품명"] = base["상품명"].fillna("").astype(str).str.strip()
    base["현재고수량"] = pd.to_numeric(
        base["현재고수량"], errors="coerce"
    ).fillna(0)

    base = base.dropna(subset=["기준일"]).copy()

    validation["base_rows"] = len(base)
    validation["base_qty"] = float(base["현재고수량"].sum())
    validation["base_date_values"] = sorted(
        base["기준일"].dt.strftime("%Y-%m-%d").unique().tolist()
    )
    validation["base_negative_rows"] = int(
        (base["현재고수량"] < 0).sum()
    )

    # 기초재고 상품코드 중복 검증
    validation["base_duplicate_skus"] = int(
        base["상품코드"].duplicated(keep=False).sum()
    )

    df = pd.concat(frames, ignore_index=True, sort=False)
    df = df.dropna(subset=["작업일"]).copy()
    df["년월"] = df["작업일"].dt.to_period("M").astype(str)

    for col in ["상품명", "상품코드", "공급처", "공급처 상품명", "전표제목"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    df["공장"] = df["공급처 상품명"].apply(classify_factory)
    df["불량타입"] = df["전표제목"].apply(classify_defect_type)

    return df, base, validation


def safe_rate(numerator: float, denominator: float):
    if denominator == 0:
        return 0.0
    return numerator / denominator * 100.0


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
    mask = (
        df["공장"].isin(selected_factories)
        & df["출입구분"].isin(selected_kinds)
        & (df["작업일"].dt.date >= start_date)
        & (df["작업일"].dt.date <= end_date)
    )
    result = df.loc[mask].copy()

    # 핵심: 불량타입 필터는 불량 행에만 적용.
    if not result.empty:
        defect_mask = result["출입구분"].eq("불량")
        result = result[
            (~defect_mask)
            | result["불량타입"].isin(selected_defect_types)
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


def filtered_base_stock(
    base: pd.DataFrame,
    filtered_transactions: pd.DataFrame,
) -> pd.DataFrame:
    """
    기초재고에는 공장/불량타입 정보가 없으므로,
    선택된 거래 데이터에 등장하는 SKU만 기초재고를 연결한다.
    상품코드가 선택 조건에 맞는 SKU의 기초재고를 사용한다.
    """
    if filtered_transactions.empty:
        return base.iloc[0:0].copy()

    sku_set = set(filtered_transactions["상품코드"].astype(str))
    return base[base["상품코드"].isin(sku_set)].copy()


def monthly_inventory_summary(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
) -> pd.DataFrame:
    """
    월별 재고 흐름:
    월초 이론재고
    + 당월 입고
    = 당월 가용재고
    - 당월 불량
    = 월말 이론재고
    """
    if filtered.empty:
        return pd.DataFrame(
            columns=[
                "년월", "기초재고", "입고", "불량",
                "월초이론재고", "당월가용재고", "월말이론재고",
                "월불량률", "순수입고대비불량률",
                "누적입고", "누적불량", "누계불량률",
            ]
        )

    base_sub = filtered_base_stock(base, filtered)
    base_qty = float(base_sub["현재고수량"].sum())

    monthly = (
        filtered.groupby(["년월", "출입구분"])["수량"]
        .sum()
        .unstack(fill_value=0)
        .sort_index()
    )

    for col in DATA_TYPES_ORDER:
        if col not in monthly.columns:
            monthly[col] = 0.0

    monthly = monthly[DATA_TYPES_ORDER].copy()
    monthly["입고"] = pd.to_numeric(monthly["입고"], errors="coerce").fillna(0)
    monthly["불량"] = pd.to_numeric(monthly["불량"], errors="coerce").fillna(0)

    monthly["기초재고"] = base_qty
    monthly["누적입고"] = monthly["입고"].cumsum()
    monthly["누적불량"] = monthly["불량"].cumsum()

    monthly["월초이론재고"] = (
        base_qty
        + monthly["누적입고"].shift(1, fill_value=0)
        - monthly["누적불량"].shift(1, fill_value=0)
    )
    monthly["당월가용재고"] = monthly["월초이론재고"] + monthly["입고"]
    monthly["월말이론재고"] = (
        monthly["당월가용재고"] - monthly["불량"]
    )

    # 기본 불량률: 기초재고까지 포함한 가용재고 기준
    monthly["월불량률"] = (
        monthly["불량"]
        .div(monthly["당월가용재고"].replace(0, pd.NA))
        * 100
    ).fillna(0)

    # 참고용: 순수하게 해당 기간 입고량만 분모로 한 기존 방식
    monthly["순수입고대비불량률"] = (
        monthly["불량"]
        .div(monthly["입고"].replace(0, pd.NA))
        * 100
    ).fillna(0)

    # 누계 불량률 역시 기초재고를 포함
    누계가용 = base_qty + monthly["누적입고"]
    monthly["누계불량률"] = (
        monthly["누적불량"]
        .div(누계가용.replace(0, pd.NA))
        * 100
    ).fillna(0)

    monthly.index.name = "년월"
    return monthly.reset_index()


def sku_risk_table(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
    min_available_qty: int,
    top_n: int,
) -> pd.DataFrame:
    """
    SKU 위험도:
    가용량 = 기초재고 + 누적 입고
    불량률 = 누적 불량 / 가용량
    월말이론재고 = 가용량 - 누적 불량
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

    base_group = (
        base.groupby("상품코드", as_index=False)
        .agg(기초재고=("현재고수량", "sum"))
    )

    # 입고만 존재하는 SKU뿐 아니라 불량만 있는 SKU도 잡기 위해
    # 거래 전체 SKU를 기준으로 만든다.
    all_skus = pd.DataFrame(
        {"상품코드": sorted(set(filtered["상품코드"].astype(str)))}
    )

    result = all_skus.merge(base_group, on="상품코드", how="left")
    result = result.merge(inbound_group, on="상품코드", how="left")
    result = result.merge(defect_group, on="상품코드", how="left")

    result["기초재고"] = result["기초재고"].fillna(0)
    result["입고수량"] = result["입고수량"].fillna(0)
    result["불량수량"] = result["불량수량"].fillna(0)
    result["상품명"] = result["상품명"].fillna("")
    result["공장"] = result["공장"].fillna("기초재고만/거래공장 미상")

    result["가용량"] = result["기초재고"] + result["입고수량"]
    result["월말이론재고"] = result["가용량"] - result["불량수량"]

    result = result[result["가용량"] >= min_available_qty].copy()
    if result.empty:
        return result

    result["불량률"] = (
        result["불량수량"]
        .div(result["가용량"].replace(0, pd.NA))
        * 100
    ).fillna(0)

    result["재고흐름"] = result["월말이론재고"].apply(
        lambda x: "점검필요(음수)" if x < 0 else "정상범위"
    )

    result = result.sort_values(
        ["재고흐름", "불량률", "불량수량", "가용량"],
        ascending=[True, False, False, False],
        key=lambda s: (
            s if s.name != "재고흐름"
            else s.map({"점검필요(음수)": 0, "정상범위": 1})
        ),
    ).head(top_n)

    result["순위"] = range(1, len(result) + 1)

    return result[
        [
            "순위", "상품코드", "상품명", "공장",
            "기초재고", "입고수량", "가용량",
            "불량수량", "불량률", "월말이론재고", "재고흐름",
        ]
    ].reset_index(drop=True)


def stock_flow_validation(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
) -> pd.DataFrame:
    """SKU별 기초 + 입고 - 불량 흐름을 검증한다."""
    if filtered.empty:
        return pd.DataFrame()

    tx = filtered.copy()
    base_sub = filtered_base_stock(base, tx)

    base_group = (
        base_sub.groupby("상품코드", as_index=False)
        .agg(기초재고=("현재고수량", "sum"))
    )
    in_group = (
        tx[tx["출입구분"] == "입고"]
        .groupby("상품코드", as_index=False)
        .agg(입고수량=("수량", "sum"))
    )
    def_group = (
        tx[tx["출입구분"] == "불량"]
        .groupby("상품코드", as_index=False)
        .agg(불량수량=("수량", "sum"))
    )

    sku = pd.DataFrame(
        {"상품코드": sorted(set(tx["상품코드"].astype(str)))}
    )
    sku = sku.merge(base_group, on="상품코드", how="left")
    sku = sku.merge(in_group, on="상품코드", how="left")
    sku = sku.merge(def_group, on="상품코드", how="left")

    for c in ["기초재고", "입고수량", "불량수량"]:
        sku[c] = sku[c].fillna(0)

    sku["가용량"] = sku["기초재고"] + sku["입고수량"]
    sku["이론재고"] = sku["가용량"] - sku["불량수량"]
    sku["상태"] = sku["이론재고"].apply(
        lambda x: "점검필요" if x < 0 else "정상"
    )
    return sku.sort_values(
        ["상태", "이론재고"],
        ascending=[True, True],
        key=lambda s: (
            s if s.name != "상태"
            else s.map({"점검필요": 0, "정상": 1})
        ),
    ).reset_index(drop=True)


# ------------------------------------------------------------------
# 화면 시작
# ------------------------------------------------------------------
st.title("📦 입고 / 불량 / 기초재고 대시보드")
st.caption(
    "2026-01-01 기초재고를 출발점으로 입고·불량 재고 흐름과 "
    "불량률을 계산합니다."
)

data_source = None
if os.path.exists(DEFAULT_DATA_PATH):
    data_source = DEFAULT_DATA_PATH

uploaded = st.file_uploader(
    "raw_data.xlsx 업로드 (선택)",
    type=["xlsx"],
    help="업로드하면 업로드한 파일을 사용합니다.",
)
if uploaded is not None:
    data_source = io.BytesIO(uploaded.getvalue())

if data_source is None:
    st.info("data/raw_data.xlsx를 넣거나 위에서 엑셀 파일을 업로드해주세요.")
    st.stop()

try:
    df, base, validation = load_and_process(data_source)
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()


# ------------------------------------------------------------------
# 검증
# ------------------------------------------------------------------
with st.expander("🔎 원본 데이터 / 기초재고 검증", expanded=True):
    a, b, c, d, e = st.columns(5)
    a.metric("기초재고", f"{validation['base_qty']:,.0f}")
    b.metric("총 입고", f"{validation['source_qty']['입고']:,.0f}")
    c.metric("총 불량", f"{validation['source_qty']['불량']:,.0f}")
    d.metric("기초재고 SKU", f"{base['상품코드'].nunique():,}")
    e.metric("기초재고 기준일", ", ".join(validation["base_date_values"]))

    st.write(
        "원본 행 수:",
        {
            "입고": validation["source_rows"]["입고"],
            "불량": validation["source_rows"]["불량"],
        },
    )
    st.write(
        "음수 수량 행:",
        {
            "입고": validation["negative_qty_rows"]["입고"],
            "불량": validation["negative_qty_rows"]["불량"],
            "기초재고": validation["base_negative_rows"],
        },
    )
    st.write(
        "날짜 오류 행:",
        validation["invalid_dates"],
    )

    if len(validation["base_date_values"]) == 1 and \
       validation["base_date_values"][0] == "2026-01-01":
        st.success("기초재고 기준일은 2026-01-01로 확인되었습니다.")
    else:
        st.warning(
            "기초재고 기준일이 2026-01-01 하나로만 구성되어 있지 않습니다. "
            "현재 파일의 실제 기준일을 확인해주세요."
        )

    defect_check = (
        df[df["출입구분"] == "불량"]
        .groupby("불량타입")["수량"]
        .agg(["count", "sum"])
        .reindex(DEFECT_TYPES_ORDER)
        .fillna(0)
    )
    defect_check = defect_check.rename(
        columns={"count": "행수", "sum": "수량"}
    )
    st.dataframe(
        defect_check.style.format(
            {"행수": "{:,.0f}", "수량": "{:,.0f}"}
        ),
        use_container_width=True,
    )

    st.caption(
        "※ 불량타입 필터는 불량 행에만 적용됩니다. "
        "기초재고에는 공장/불량타입 정보가 없으므로 SKU 기준으로 재고 흐름에 연결합니다."
    )


# ------------------------------------------------------------------
# 필터 - Excel 필터처럼 전체 체크 상태에서 해제
# ------------------------------------------------------------------
st.subheader("🔎 필터")

filter1, filter2, filter3 = st.columns(3)

factory_options = sorted(df["공장"].dropna().unique().tolist())
defect_type_options = [
    x for x in DEFECT_TYPES_ORDER if x in set(df["불량타입"])
]

with filter1:
    selected_factories = st.multiselect(
        "🏭 공장",
        options=factory_options,
        default=factory_options,
        help="Excel 필터처럼 처음에는 전체 선택됩니다. 체크를 해제해 제외하세요.",
    )

with filter2:
    selected_defect_types = st.multiselect(
        "🏷️ 불량타입",
        options=defect_type_options,
        default=defect_type_options,
        help="처음에는 전체 선택. 체크 해제한 불량타입만 제외됩니다.",
    )

with filter3:
    selected_kinds = st.multiselect(
        "📂 데이터 종류",
        options=DATA_TYPES_ORDER,
        default=DATA_TYPES_ORDER,
        help="처음에는 입고/불량 전체 선택. 체크 해제로 원하는 종류만 남길 수 있습니다.",
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
    min_date = df["작업일"].min().date()
    max_date = df["작업일"].max().date()
    date_range = st.date_input(
        "작업일 범위",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

if not selected_factories:
    st.warning("공장을 하나 이상 선택해주세요.")
    st.stop()

if not selected_kinds:
    st.warning("데이터 종류를 하나 이상 선택해주세요.")
    st.stop()

if not selected_defect_types and "불량" in selected_kinds:
    st.warning("불량 데이터를 보려면 불량타입을 하나 이상 선택해주세요.")
    st.stop()

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

base_filtered = filtered_base_stock(base, filtered)
total_base = base_filtered["현재고수량"].sum()
available_total = total_base + total_in
ending_theoretical = available_total - total_defect
overall_rate = safe_rate(total_defect, available_total)
pure_inbound_rate = safe_rate(total_defect, total_in)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("선택 기초재고", f"{total_base:,.0f}")
k2.metric("선택 입고", f"{total_in:,.0f}")
k3.metric("선택 불량", f"{total_defect:,.0f}")
k4.metric("가용량", f"{available_total:,.0f}")
k5.metric("가용재고 대비 불량률", f"{overall_rate:.2f}%")
k6.metric("이론 잔여재고", f"{ending_theoretical:,.0f}")

if ending_theoretical < 0:
    st.warning(
        "선택 조건에서 기초재고 + 입고 - 불량이 음수입니다. "
        "이 파일에 없는 판매/출고/이동 등의 재고 감소 거래가 있다면 "
        "실제 재고 흐름과 다를 수 있으므로 아래 검증표를 확인하세요."
    )

st.divider()


# ------------------------------------------------------------------
# 월별 재고 흐름
# ------------------------------------------------------------------
monthly = monthly_inventory_summary(filtered, base)

st.subheader("📊 월별 재고 흐름 / 불량률")

fig_month = go.Figure()
_, defect_colors = chart_colors()

fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["입고"],
        name="입고",
        text=monthly["입고"].map(lambda x: f"{x:,.0f}"),
        textposition="outside",
        marker_color="#42A5F5" if get_theme_type() == "dark" else "#1565C0",
        opacity=0.78,
        hovertemplate="입고: %{y:,.0f}<extra></extra>",
    )
)
fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["불량"],
        name="불량",
        text=monthly["불량"].map(lambda x: f"{x:,.0f}"),
        textposition="outside",
        marker_color="#EF5350" if get_theme_type() == "dark" else "#C62828",
        opacity=0.86,
        hovertemplate="불량: %{y:,.0f}<extra></extra>",
    )
)
fig_month.add_trace(
    go.Scatter(
        x=monthly["년월"],
        y=monthly["월불량률"],
        name="가용재고 대비 월불량률",
        mode="lines+markers+text",
        text=monthly["월불량률"].map(lambda x: f"{x:.2f}%"),
        textposition="top center",
        line=dict(
            color="#FFD54F" if get_theme_type() == "dark" else "#E65100",
            width=3,
        ),
        marker=dict(size=8),
        yaxis="y2",
        hovertemplate="가용재고 대비 불량률: %{y:.2f}%<extra></extra>",
    )
)
fig_month = apply_theme(fig_month, 540)
fig_month.update_layout(
    barmode="group",
    xaxis_title="년월",
    yaxis=dict(title="수량", separatethousands=True),
    yaxis2=dict(
        title="불량률 (%)",
        overlaying="y",
        side="right",
        rangemode="tozero",
        ticksuffix="%",
        showgrid=False,
    ),
)
st.plotly_chart(fig_month, use_container_width=True)

st.caption(
    "핵심 기준: 월불량률 = 당월 불량 ÷ (월초 이론재고 + 당월 입고). "
    "기초재고가 존재하므로 단순히 '불량 ÷ 입고'로 계산하지 않습니다."
)


# ------------------------------------------------------------------
# 월별 누계
# ------------------------------------------------------------------
st.subheader("📈 월별 누계")

tab1, tab2, tab3, tab4 = st.tabs(
    ["누계 입고", "누계 불량", "누계 이론재고", "누계 불량률"]
)

def line_chart(x, y, name, title, y_title, suffix="", color_light="#1565C0", color_dark="#64B5F6"):
    fig = go.Figure()
    color = color_dark if get_theme_type() == "dark" else color_light
    fig.add_trace(
        go.Scatter(
            x=x,
            y=y,
            name=name,
            mode="lines+markers+text",
            text=[
                f"{v:,.2f}{suffix}" if suffix else f"{v:,.0f}"
                for v in y
            ],
            textposition="top center",
            line=dict(color=color, width=4),
            marker=dict(size=8),
        )
    )
    fig = apply_theme(fig, 420)
    fig.update_layout(
        title=title,
        xaxis_title="년월",
        yaxis_title=y_title,
        yaxis=dict(ticksuffix=suffix),
    )
    return fig

with tab1:
    st.plotly_chart(
        line_chart(
            monthly["년월"],
            monthly["누적입고"],
            "누계 입고",
            "누계 입고 수량",
            "수량",
            color_light="#1565C0",
            color_dark="#64B5F6",
        ),
        use_container_width=True,
    )

with tab2:
    st.plotly_chart(
        line_chart(
            monthly["년월"],
            monthly["누적불량"],
            "누계 불량",
            "누계 불량 수량",
            "수량",
            color_light="#C62828",
            color_dark="#EF5350",
        ),
        use_container_width=True,
    )

with tab3:
    st.plotly_chart(
        line_chart(
            monthly["년월"],
            monthly["월말이론재고"],
            "월말 이론재고",
            "기초재고 + 누계 입고 - 누계 불량",
            "이론재고",
            color_light="#2E7D32",
            color_dark="#81C784",
        ),
        use_container_width=True,
    )

with tab4:
    st.plotly_chart(
        line_chart(
            monthly["년월"],
            monthly["누계불량률"],
            "누계 가용재고 대비 불량률",
            "누계 가용재고 대비 불량률",
            "불량률",
            suffix="%",
            color_light="#E65100",
            color_dark="#FFD54F",
        ),
        use_container_width=True,
    )


# ------------------------------------------------------------------
# 월별 숫자 검증표
# ------------------------------------------------------------------
with st.expander("📋 월별 계산 검증표", expanded=False):
    show_month = monthly[
        [
            "년월", "기초재고", "월초이론재고", "입고", "당월가용재고",
            "불량", "월말이론재고", "월불량률",
            "누적입고", "누적불량", "누계불량률",
            "순수입고대비불량률",
        ]
    ].copy()

    st.dataframe(
        show_month.style.format(
            {
                "기초재고": "{:,.0f}",
                "월초이론재고": "{:,.0f}",
                "입고": "{:,.0f}",
                "당월가용재고": "{:,.0f}",
                "불량": "{:,.0f}",
                "월말이론재고": "{:,.0f}",
                "월불량률": "{:.2f}%",
                "누적입고": "{:,.0f}",
                "누적불량": "{:,.0f}",
                "누계불량률": "{:.2f}%",
                "순수입고대비불량률": "{:.2f}%",
            }
        ),
        use_container_width=True,
    )


# ------------------------------------------------------------------
# 불량타입
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

    def make_defect_group_chart(data, types, title):
        _, colors = chart_colors()
        fig = go.Figure()
        for defect_type in types:
            fig.add_trace(
                go.Bar(
                    x=data.index,
                    y=data[defect_type],
                    name=defect_type,
                    marker_color=colors[defect_type],
                    text=data[defect_type].map(
                        lambda x: f"{x:,.0f}" if x != 0 else ""
                    ),
                    textposition="outside",
                    hovertemplate=(
                        f"{defect_type}: %{y:,.0f}<extra></extra>"
                    ),
                )
            )
        fig = apply_theme(fig, 500)
        fig.update_layout(
            title=title,
            barmode="group",
            xaxis_title="년월",
            yaxis_title="불량 수량",
        )
        return fig

    with tab_main:
        main_types = [
            x for x in ["테", "렌즈"]
            if type_monthly[x].abs().sum() != 0
        ]
        if main_types:
            st.plotly_chart(
                make_defect_group_chart(
                    type_monthly, main_types, "테 / 렌즈 월별 불량 수량"
                ),
                use_container_width=True,
            )
        else:
            st.info("테/렌즈 데이터가 없습니다.")

    with tab_minor:
        minor_types = [
            x for x in ["전체", "기타"]
            if type_monthly[x].abs().sum() != 0
        ]
        if minor_types:
            st.plotly_chart(
                make_defect_group_chart(
                    type_monthly, minor_types, "전체 / 기타 월별 불량 수량"
                ),
                use_container_width=True,
            )
        else:
            st.info("전체/기타 데이터가 없습니다.")

    with tab_table:
        check_table = type_monthly.copy()
        check_table["불량 합계"] = check_table.sum(axis=1)
        st.dataframe(
            check_table.style.format("{:,.0f}"),
            use_container_width=True,
        )
        st.caption(
            "그룹형 막대그래프이므로 각 타입이 같은 기준선에서 직접 비교됩니다."
        )


# ------------------------------------------------------------------
# 공장별
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
factory_summary = factory_summary[DATA_TYPES_ORDER]

# 공장별 기초재고는 기초재고에 공장 정보가 없으므로
# 거래에 등장하는 SKU를 기준으로 연결.
factory_base_rows = []
for factory in factory_summary.index:
    tx_factory = filtered[filtered["공장"] == factory]
    b = filtered_base_stock(base, tx_factory)
    factory_base_rows.append(
        {
            "공장": factory,
            "기초재고": b["현재고수량"].sum(),
        }
    )
factory_base = pd.DataFrame(factory_base_rows).set_index("공장")

factory_summary["기초재고"] = factory_base["기초재고"].reindex(
    factory_summary.index
).fillna(0)
factory_summary["가용량"] = (
    factory_summary["기초재고"] + factory_summary["입고"]
)
factory_summary["불량률"] = (
    factory_summary["불량"]
    .div(factory_summary["가용량"].replace(0, pd.NA))
    * 100
).fillna(0)
factory_summary["이론잔여재고"] = (
    factory_summary["가용량"] - factory_summary["불량"]
)

factory_colors, _ = chart_colors()

fig_factory = go.Figure()
fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["입고"],
        name="입고",
        marker_color=[
            factory_colors.get(x, factory_colors["미상"])
            for x in factory_summary.index
        ],
    )
)
fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["불량"],
        name="불량",
        marker_color=(
            "#EF5350" if get_theme_type() == "dark" else "#C62828"
        ),
    )
)
fig_factory.add_trace(
    go.Scatter(
        x=factory_summary.index,
        y=factory_summary["불량률"],
        name="가용량 대비 불량률",
        mode="lines+markers+text",
        text=factory_summary["불량률"].map(lambda x: f"{x:.2f}%"),
        textposition="top center",
        yaxis="y2",
        line=dict(
            color="#FFD54F" if get_theme_type() == "dark" else "#E65100",
            width=3,
        ),
        marker=dict(size=9),
    )
)
fig_factory = apply_theme(fig_factory, 480)
fig_factory.update_layout(
    barmode="group",
    xaxis_title="공장",
    yaxis=dict(title="수량"),
    yaxis2=dict(
        title="불량률 (%)",
        overlaying="y",
        side="right",
        rangemode="tozero",
        ticksuffix="%",
        showgrid=False,
    ),
)
st.plotly_chart(fig_factory, use_container_width=True)

st.dataframe(
    factory_summary.style.format(
        {
            "기초재고": "{:,.0f}",
            "입고": "{:,.0f}",
            "불량": "{:,.0f}",
            "가용량": "{:,.0f}",
            "불량률": "{:.2f}%",
            "이론잔여재고": "{:,.0f}",
        }
    ),
    use_container_width=True,
)


# ------------------------------------------------------------------
# SKU 위험도
# ------------------------------------------------------------------
st.divider()
st.subheader("🚨 기초재고 + 입고 기준 SKU별 불량 위험")

s1, s2 = st.columns([1, 1])
with s1:
    min_available_qty = st.number_input(
        "최소 가용량",
        min_value=1,
        value=100,
        step=10,
        help=(
            "기초재고 + 누적 입고가 너무 작은 SKU는 작은 불량 몇 개만으로 "
            "불량률이 과도하게 커질 수 있어 최소 기준을 둡니다."
        ),
    )
with s2:
    top_n = st.number_input(
        "표시 SKU 수",
        min_value=5,
        max_value=100,
        value=30,
        step=5,
    )

sku_table = sku_risk_table(
    filtered=filtered,
    base=base,
    min_available_qty=min_available_qty,
    top_n=top_n,
)

if sku_table.empty:
    st.info("현재 조건에서 최소 가용량 조건을 만족하는 SKU가 없습니다.")
else:
    st.dataframe(
        sku_table.style.format(
            {
                "기초재고": "{:,.0f}",
                "입고수량": "{:,.0f}",
                "가용량": "{:,.0f}",
                "불량수량": "{:,.0f}",
                "불량률": "{:.2f}%",
                "월말이론재고": "{:,.0f}",
            }
        ),
        use_container_width=True,
    )
    st.caption(
        "불량률 = 누적 불량 ÷ (기초재고 + 누적 입고). "
        "'점검필요(음수)'는 해당 SKU의 기초재고 + 입고보다 불량이 많다는 뜻입니다."
    )


# ------------------------------------------------------------------
# 재고 흐름 이상 SKU
# ------------------------------------------------------------------
st.divider()
st.subheader("⚠️ 재고 흐름 점검")

flow_check = stock_flow_validation(filtered, base)

if flow_check.empty:
    st.info("검증할 SKU가 없습니다.")
else:
    issue = flow_check[flow_check["상태"] == "점검필요"].copy()
    f1, f2, f3 = st.columns(3)
    f1.metric("검증 SKU", f"{len(flow_check):,}")
    f2.metric("음수 이론재고 SKU", f"{len(issue):,}")
    f3.metric(
        "음수 이론재고 합계",
        f"{issue['이론재고'].sum():,.0f}" if not issue.empty else "0",
    )

    if issue.empty:
        st.success(
            "현재 선택 조건에서는 기초재고 + 입고 - 불량이 음수가 되는 SKU가 없습니다."
        )
    else:
        st.warning(
            "아래 SKU는 기초재고 + 입고만으로는 기록된 불량량을 설명할 수 없습니다. "
            "판매/출고/이동 등 다른 재고 감소 거래가 원본에 없는지 확인해야 합니다."
        )
        st.dataframe(
            issue.head(100).style.format(
                {
                    "기초재고": "{:,.0f}",
                    "입고수량": "{:,.0f}",
                    "불량수량": "{:,.0f}",
                    "가용량": "{:,.0f}",
                    "이론재고": "{:,.0f}",
                }
            ),
            use_container_width=True,
        )


# ------------------------------------------------------------------
# 상세 데이터
# ------------------------------------------------------------------
st.divider()
st.subheader("📄 필터링된 원본 데이터")

display_cols = [
    c for c in [
        "작업일", "출입구분", "공장", "불량타입",
        "상품코드", "상품명", "수량", "전표제목",
        "공급처 상품명", "현재고", "작업자", "전표번호"
    ] if c in filtered.columns
]

st.dataframe(
    filtered[display_cols].sort_values(
        ["작업일", "출입구분", "상품코드"],
        ascending=[True, True, True],
    ),
    use_container_width=True,
    height=500,
)

csv = filtered[display_cols].to_csv(index=False, encoding="utf-8-sig")
st.download_button(
    "⬇️ 현재 필터 데이터 CSV 다운로드",
    data=csv,
    file_name="입고_불량_필터데이터.csv",
    mime="text/csv",
)
