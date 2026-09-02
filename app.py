# -*- coding: utf-8 -*-
"""
입고 / 출고 / 불량 / 기초재고 / 상품마스터 대시보드
=================================================

재고 및 검수 기준
1) 2026-01-01 기초재고를 출발점으로 사용한다.
2) 재고 흐름 = 기초재고 + 누적 입고 - 누적 출고 - 누적 불량
3) 출고를 맞추기 위한 검수 과정에서 정상 출고수량과 불량수량이 발생한다.
4) 실제 검수수량 = 출고수량 + 불량수량
5) 검수 기준 불량률 = 불량수량 / (출고수량 + 불량수량)
6) 아직 검수하지 않은 잔여재고는 불량률의 분모에 포함하지 않는다.
7) 상품마스터의 상품코드를 기준으로 공장명(C2/C2-S/C5/미상),
   카테고리(G=안경/S=선글라스)를 연결한다.
8) 상품마스터가 공장/카테고리의 기준정보이며,
   마스터 미등록 SKU는 미상으로 표시한다.
9) 기초재고는 상품코드로 상품마스터와 연결하여 공장/카테고리 필터에 반영한다.
10) 음수 이론재고는 재고흐름 점검 대상으로 별도 표시한다.
11) 화면의 DataFrame 인덱스는 숨긴다.
12) 검증 정보는 dict/JSON 형태가 아닌 정상적인 표/카드 형태로 표시한다.
"""

import io
import os
from typing import List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ==============================================================
# 기본 설정
# ==============================================================

st.set_page_config(
    page_title="입고 / 출고 / 불량 대시보드",
    page_icon="📦",
    layout="wide",
)

DEFAULT_DATA_PATH = os.path.join(
    os.path.dirname(__file__),
    "data",
    "raw_data.xlsx",
)

FACTORIES = [
    "C2공장",
    "C2-S공장",
    "C5공장",
    "미상",
]

CATEGORIES = [
    "G",
    "S",
    "미상",
]

DEFECT_TYPES_ORDER = [
    "테",
    "렌즈",
    "전체",
    "기타",
]

DATA_TYPES_ORDER = [
    "입고",
    "출고",
    "불량",
]

FACTORY_COLORS_LIGHT = {
    "C5공장": "#1565C0",
    "C2공장": "#2E7D32",
    "C2-S공장": "#EF6C00",
    "미상": "#757575",
}

FACTORY_COLORS_DARK = {
    "C5공장": "#64B5F6",
    "C2공장": "#81C784",
    "C2-S공장": "#FFB74D",
    "미상": "#BDBDBD",
}

REQUIRED_TRANSACTION_COLUMNS = [
    "작업일",
    "수량",
    "공급처 상품명",
    "상품코드",
    "상품명",
    "전표제목",
]

REQUIRED_BASE_COLUMNS = [
    "기준일",
    "상품코드",
    "상품명",
    "현재고수량",
]

REQUIRED_OUTBOUND_COLUMNS = [
    "출고일",
    "상품코드",
    "상품명",
    "출고수량",
]

REQUIRED_MASTER_COLUMNS = [
    "상품코드",
    "상품명",
    "공장명",
    "카테고리",
]


# ==============================================================
# 공통 화면 함수
# ==============================================================

def get_theme_type() -> str:
    """Streamlit 현재 테마를 가져온다."""
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

    return "light"


def hide_dataframe_index(styler):
    """
    DataFrame의 화면상 인덱스를 숨긴다.

    중요:
    - 데이터 자체의 index를 삭제하지 않는다.
    - 화면에서만 index가 보이지 않도록 한다.
    """
    try:
        return styler.hide(axis="index")
    except Exception:
        try:
            return styler.hide_index()
        except Exception:
            return styler


def apply_theme(
    fig: go.Figure,
    height: int = 480,
):
    """Plotly 공통 테마."""
    dark = get_theme_type() == "dark"

    template = "plotly_dark" if dark else "plotly_white"

    grid_color = (
        "rgba(255,255,255,0.15)"
        if dark
        else "rgba(0,0,0,0.10)"
    )

    zero_color = (
        "rgba(255,255,255,0.35)"
        if dark
        else "rgba(0,0,0,0.25)"
    )

    fig.update_layout(
        template=template,
        height=height,
        margin=dict(
            t=55,
            l=20,
            r=75,
            b=35,
        ),
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


# ==============================================================
# 데이터 분류
# ==============================================================

def classify_factory(value) -> str:
    if pd.isna(value):
        return "미상"

    text = str(value).strip().upper()

    if text == "":
        return "미상"

    if "C2-S" in text or "C2S" in text:
        return "C2-S공장"

    if "C2" in text:
        return "C2공장"

    if "C5" in text:
        return "C5공장"

    return "미상"


def classify_category(value) -> str:
    if pd.isna(value):
        return "미상"

    text = str(value).strip().upper()

    if text == "G":
        return "G"

    if text == "S":
        return "S"

    return "미상"


def classify_defect_type(value) -> str:
    if pd.isna(value):
        return "기타"

    text = str(value).strip()

    if text == "":
        return "기타"

    if "전체" in text:
        return "전체"

    if "테" in text:
        return "테"

    if "렌즈" in text:
        return "렌즈"

    return "기타"


def normalize_code(series: pd.Series) -> pd.Series:
    """
    상품코드의 엑셀 숫자형/문자형 차이를 최대한 흡수한다.
    예:
    12345.0 -> 12345
    """
    s = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
    )

    s = s.str.replace(
        r"\.0$",
        "",
        regex=True,
    )

    return s


# ==============================================================
# 상품마스터
# ==============================================================

def prepare_master(
    master: pd.DataFrame,
) -> Tuple[pd.DataFrame, dict]:

    missing = [
        c
        for c in REQUIRED_MASTER_COLUMNS
        if c not in master.columns
    ]

    if missing:
        raise ValueError(
            f"'상품마스터' 시트에 필수 컬럼이 없습니다: {missing}"
        )

    m = master.copy()

    m["상품코드"] = normalize_code(
        m["상품코드"]
    )

    m["상품명"] = (
        m["상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    m["공장명"] = (
        m["공장명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    m["카테고리"] = (
        m["카테고리"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # 빈 코드 제외하고 중복 확인
    duplicate_mask = (
        m["상품코드"].duplicated(keep=False)
        & m["상품코드"].ne("")
    )

    duplicate_codes = sorted(
        m.loc[
            duplicate_mask,
            "상품코드",
        ]
        .unique()
        .tolist()
    )

    # 상품코드 중복이면 첫 번째 행 사용
    m = m.drop_duplicates(
        subset=["상품코드"],
        keep="first",
    ).copy()

    m["공장"] = m["공장명"].apply(
        classify_factory
    )

    m["카테고리"] = m["카테고리"].apply(
        classify_category
    )

    validation = {
        "master_rows": len(master),
        "master_unique_skus": int(
            m.loc[
                m["상품코드"].ne(""),
                "상품코드",
            ].nunique()
        ),
        "master_duplicate_codes": duplicate_codes,
        "master_blank_codes": int(
            m["상품코드"].eq("").sum()
        ),
        "master_unknown_factory": int(
            (m["공장"] == "미상").sum()
        ),
        "master_unknown_category": int(
            (m["카테고리"] == "미상").sum()
        ),
    }

    return m, validation


# ==============================================================
# 상품마스터 연결
# ==============================================================

def attach_master(
    df: pd.DataFrame,
    master: pd.DataFrame,
    source_name: str,
) -> Tuple[pd.DataFrame, dict]:

    out = df.copy()

    out["상품코드"] = normalize_code(
        out["상품코드"]
    )

    lookup = master[
        [
            "상품코드",
            "상품명",
            "공장",
            "카테고리",
        ]
    ].copy()

    lookup = lookup.rename(
        columns={
            "상품명": "마스터상품명",
            "공장": "공장_마스터",
            "카테고리": "카테고리_마스터",
        }
    )

    out = out.merge(
        lookup,
        on="상품코드",
        how="left",
    )

    out["공장"] = (
        out["공장_마스터"]
        .fillna("미상")
    )

    out["카테고리"] = (
        out["카테고리_마스터"]
        .fillna("미상")
    )

    # 마스터 상품명을 우선
    if "상품명" in out.columns:
        out["상품명"] = (
            out["마스터상품명"]
            .where(
                out["마스터상품명"]
                .fillna("")
                .ne(""),
                out["상품명"]
                .fillna("")
                .astype(str),
            )
        )

    out = out.drop(
        columns=[
            "마스터상품명",
            "공장_마스터",
            "카테고리_마스터",
        ],
        errors="ignore",
    )

    master_code_set = set(
        master["상품코드"]
    )

    missing_master = sorted(
        out.loc[
            ~out["상품코드"].isin(master_code_set)
            & out["상품코드"].ne(""),
            "상품코드",
        ]
        .unique()
        .tolist()
    )

    return out, {
        "source": source_name,
        "rows": len(out),
        "missing_master_skus": missing_master,
    }


# ==============================================================
# 엑셀
# ==============================================================

def read_sheet(
    file_source,
    sheet_name: str,
) -> pd.DataFrame:

    return pd.read_excel(
        file_source,
        sheet_name=sheet_name,
    )


@st.cache_data(
    show_spinner="엑셀 데이터를 불러오고 검증하는 중입니다..."
)
def load_and_process(
    file_bytes: bytes,
) -> Tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    dict,
]:

    source = io.BytesIO(file_bytes)

    validation = {
        "source_rows": {},
        "source_qty": {},
        "negative_qty_rows": {},
        "invalid_dates": {},
        "base_rows": 0,
        "base_qty": 0,
        "base_date_values": [],
        "base_negative_rows": 0,
        "base_duplicate_skus": 0,
        "master": {},
        "missing_master": {},
    }

    # ----------------------------------------------------------
    # 상품마스터
    # ----------------------------------------------------------

    try:
        master_raw = read_sheet(
            source,
            "상품마스터",
        )
    except Exception as e:
        raise ValueError(
            f"'상품마스터' 시트를 읽을 수 없습니다: {e}"
        )

    master, master_validation = prepare_master(
        master_raw
    )

    validation["master"] = master_validation

    frames = []

    # ----------------------------------------------------------
    # 불량 / 입고
    # ----------------------------------------------------------

    for sheet_name, data_kind in [
        ("불량", "불량"),
        ("입고", "입고"),
    ]:

        df = read_sheet(
            source,
            sheet_name,
        )

        missing = [
            c
            for c in REQUIRED_TRANSACTION_COLUMNS
            if c not in df.columns
        ]

        if missing:
            raise ValueError(
                f"'{sheet_name}' 시트에 필수 컬럼이 없습니다: {missing}"
            )

        qty = pd.to_numeric(
            df["수량"],
            errors="coerce",
        ).fillna(0)

        work_date = pd.to_datetime(
            df["작업일"],
            errors="coerce",
        )

        validation["source_rows"][data_kind] = int(
            len(df)
        )

        validation["source_qty"][data_kind] = float(
            qty.sum()
        )

        validation["negative_qty_rows"][data_kind] = int(
            (qty < 0).sum()
        )

        validation["invalid_dates"][data_kind] = int(
            work_date.isna().sum()
        )

        df["출입구분"] = data_kind
        df["수량"] = qty
        df["작업일"] = work_date

        df["상품코드"] = normalize_code(
            df["상품코드"]
        )

        df, missing_info = attach_master(
            df,
            master,
            sheet_name,
        )

        validation["missing_master"][data_kind] = (
            missing_info["missing_master_skus"]
        )

        df["공급처 상품명"] = (
            df["공급처 상품명"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df["전표제목"] = (
            df["전표제목"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df["불량타입"] = (
            df["전표제목"]
            .apply(classify_defect_type)
        )

        frames.append(df)

    # ----------------------------------------------------------
    # 출고
    # ----------------------------------------------------------

    try:
        outbound = read_sheet(
            source,
            "출고",
        )
    except Exception as e:
        raise ValueError(
            f"'출고' 시트를 읽을 수 없습니다: {e}"
        )

    missing_out = [
        c
        for c in REQUIRED_OUTBOUND_COLUMNS
        if c not in outbound.columns
    ]

    if missing_out:
        raise ValueError(
            f"'출고' 시트에 필수 컬럼이 없습니다: {missing_out}"
        )

    outbound["출고일"] = pd.to_datetime(
        outbound["출고일"],
        errors="coerce",
    )

    outbound["출고수량"] = pd.to_numeric(
        outbound["출고수량"],
        errors="coerce",
    ).fillna(0)

    outbound["상품코드"] = normalize_code(
        outbound["상품코드"]
    )

    outbound, missing_info = attach_master(
        outbound,
        master,
        "출고",
    )

    validation["missing_master"]["출고"] = (
        missing_info["missing_master_skus"]
    )

    validation["source_rows"]["출고"] = int(
        len(outbound)
    )

    validation["source_qty"]["출고"] = float(
        outbound["출고수량"].sum()
    )

    validation["negative_qty_rows"]["출고"] = int(
        (outbound["출고수량"] < 0).sum()
    )

    validation["invalid_dates"]["출고"] = int(
        outbound["출고일"].isna().sum()
    )

    # 출고를 거래 구조로 변환
    outbound_tx = outbound.copy()

    outbound_tx["작업일"] = (
        outbound_tx["출고일"]
    )

    outbound_tx["수량"] = (
        outbound_tx["출고수량"]
    )

    outbound_tx["출입구분"] = "출고"

    outbound_tx["년월"] = (
        outbound_tx["작업일"]
        .dt.to_period("M")
        .astype(str)
    )

    outbound_tx["불량타입"] = "기타"
    outbound_tx["전표제목"] = ""
    outbound_tx["공급처 상품명"] = ""

    frames.append(outbound_tx)

    # ----------------------------------------------------------
    # 기초재고
    # ----------------------------------------------------------

    try:
        base = read_sheet(
            source,
            "기초재고",
        )
    except Exception as e:
        raise ValueError(
            "이번 버전은 반드시 '기초재고' 시트가 필요합니다. "
            f"기초재고 시트를 읽을 수 없습니다: {e}"
        )

    missing_base = [
        c
        for c in REQUIRED_BASE_COLUMNS
        if c not in base.columns
    ]

    if missing_base:
        raise ValueError(
            f"'기초재고' 시트에 필수 컬럼이 없습니다: {missing_base}"
        )

    base["기준일"] = pd.to_datetime(
        base["기준일"],
        errors="coerce",
    )

    base["상품코드"] = normalize_code(
        base["상품코드"]
    )

    base["상품명"] = (
        base["상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    base["현재고수량"] = pd.to_numeric(
        base["현재고수량"],
        errors="coerce",
    ).fillna(0)

    # 날짜 오류는 검증에 반영하기 위해
    # dropna 전에 계산
    base_invalid_dates = int(
        base["기준일"].isna().sum()
    )

    validation["invalid_dates"]["기초재고"] = (
        base_invalid_dates
    )

    base = base.dropna(
        subset=["기준일"]
    ).copy()

    base, missing_info = attach_master(
        base,
        master,
        "기초재고",
    )

    validation["missing_master"]["기초재고"] = (
        missing_info["missing_master_skus"]
    )

    validation["base_rows"] = int(
        len(base)
    )

    validation["base_qty"] = float(
        base["현재고수량"].sum()
    )

    validation["base_date_values"] = sorted(
        base["기준일"]
        .dt.strftime("%Y-%m-%d")
        .unique()
        .tolist()
    )

    validation["base_negative_rows"] = int(
        (base["현재고수량"] < 0).sum()
    )

    validation["base_duplicate_skus"] = int(
        base["상품코드"]
        .duplicated(keep=False)
        .sum()
    )

    # ----------------------------------------------------------
    # 거래 데이터 통합
    # ----------------------------------------------------------

    df = pd.concat(
        frames,
        ignore_index=True,
        sort=False,
    )

    df = df.dropna(
        subset=["작업일"]
    ).copy()

    df["년월"] = (
        df["작업일"]
        .dt.to_period("M")
        .astype(str)
    )

    for col in [
        "상품명",
        "상품코드",
        "공급처",
        "공급처 상품명",
        "전표제목",
        "공장",
        "카테고리",
    ]:

        if col in df.columns:
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    return (
        df,
        base,
        master,
        validation,
    )


# ==============================================================
# 계산
# ==============================================================

def safe_rate(
    numerator: float,
    denominator: float,
) -> float:

    if denominator == 0:
        return 0.0

    return (
        numerator
        / denominator
        * 100.0
    )


# ==============================================================
# 필터
# ==============================================================

def filter_data(
    df: pd.DataFrame,
    selected_factories: List[str],
    selected_categories: List[str],
    selected_defect_types: List[str],
    selected_kinds: List[str],
    start_date,
    end_date,
    product_name_query: str,
    product_code_query: str,
) -> pd.DataFrame:

    mask = (
        df["공장"].isin(selected_factories)
        & df["카테고리"].isin(selected_categories)
        & df["출입구분"].isin(selected_kinds)
        & (
            df["작업일"].dt.date
            >= start_date
        )
        & (
            df["작업일"].dt.date
            <= end_date
        )
    )

    result = df.loc[
        mask
    ].copy()

    if not result.empty:

        defect_mask = result[
            "출입구분"
        ].eq("불량")

        result = result[
            (~defect_mask)
            | result["불량타입"].isin(
                selected_defect_types
            )
        ].copy()

    if product_name_query:

        result = result[
            result["상품명"].str.contains(
                product_name_query,
                case=False,
                na=False,
                regex=False,
            )
        ].copy()

    if product_code_query:

        result = result[
            result["상품코드"].str.contains(
                product_code_query,
                case=False,
                na=False,
                regex=False,
            )
        ].copy()

    return result


def filtered_base_stock(
    base: pd.DataFrame,
    selected_factories: List[str],
    selected_categories: List[str],
    product_name_query: str = "",
    product_code_query: str = "",
) -> pd.DataFrame:

    b = base[
        base["공장"].isin(
            selected_factories
        )
        & base["카테고리"].isin(
            selected_categories
        )
    ].copy()

    if product_name_query:

        b = b[
            b["상품명"].str.contains(
                product_name_query,
                case=False,
                na=False,
                regex=False,
            )
        ].copy()

    if product_code_query:

        b = b[
            b["상품코드"].str.contains(
                product_code_query,
                case=False,
                na=False,
                regex=False,
            )
        ].copy()

    return b


# ==============================================================
# 월별 재고
# ==============================================================

def monthly_inventory_summary(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
    selected_factories: List[str],
    selected_categories: List[str],
    product_name_query: str,
    product_code_query: str,
) -> pd.DataFrame:

    columns = [
        "년월",
        "기초재고",
        "입고",
        "출고",
        "불량",
        "월초이론재고",
        "당월입고후재고",
        "월말이론재고",
        "실제검수수량",
        "월불량률",
        "누적입고",
        "누적출고",
        "누적불량",
        "누계검수수량",
        "누계불량률",
    ]

    if filtered.empty:
        return pd.DataFrame(
            columns=columns
        )

    base_sub = filtered_base_stock(
        base,
        selected_factories,
        selected_categories,
        product_name_query,
        product_code_query,
    )

    base_qty = float(
        base_sub["현재고수량"].sum()
    )

    monthly = (
        filtered
        .groupby(
            ["년월", "출입구분"]
        )["수량"]
        .sum()
        .unstack(fill_value=0)
        .sort_index()
    )

    for col in DATA_TYPES_ORDER:

        if col not in monthly.columns:
            monthly[col] = 0.0

    monthly = (
        monthly[
            DATA_TYPES_ORDER
        ]
        .copy()
        .fillna(0)
    )

    monthly["기초재고"] = base_qty

    monthly["누적입고"] = (
        monthly["입고"]
        .cumsum()
    )

    monthly["누적출고"] = (
        monthly["출고"]
        .cumsum()
    )

    monthly["누적불량"] = (
        monthly["불량"]
        .cumsum()
    )

    monthly["월초이론재고"] = (
        base_qty
        + monthly["누적입고"]
            .shift(
                1,
                fill_value=0,
            )
        - monthly["누적출고"]
            .shift(
                1,
                fill_value=0,
            )
        - monthly["누적불량"]
            .shift(
                1,
                fill_value=0,
            )
    )

    monthly["당월입고후재고"] = (
        monthly["월초이론재고"]
        + monthly["입고"]
    )

    monthly["월말이론재고"] = (
        monthly["당월입고후재고"]
        - monthly["출고"]
        - monthly["불량"]
    )

    # ----------------------------------------------------------
    # 실제 검수
    # ----------------------------------------------------------
    monthly["실제검수수량"] = (
        monthly["출고"]
        + monthly["불량"]
    )

    monthly["월불량률"] = (
        monthly["불량"]
        .div(
            monthly["실제검수수량"]
            .replace(0, pd.NA)
        )
        * 100
    ).fillna(0)

    monthly["누계검수수량"] = (
        monthly["누적출고"]
        + monthly["누적불량"]
    )

    monthly["누계불량률"] = (
        monthly["누적불량"]
        .div(
            monthly["누계검수수량"]
            .replace(0, pd.NA)
        )
        * 100
    ).fillna(0)

    monthly.index.name = "년월"

    return (
        monthly
        .reset_index()
        .copy()
    )


# ==============================================================
# SKU 요약
# ==============================================================

def build_sku_summary(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
) -> pd.DataFrame:

    if filtered.empty:
        return pd.DataFrame()

    sku_set = sorted(
        set(
            filtered["상품코드"]
            .astype(str)
        )
        - {""}
    )

    if not sku_set:
        return pd.DataFrame()

    sku = pd.DataFrame(
        {
            "상품코드": sku_set
        }
    )

    b = (
        base[
            base["상품코드"].isin(
                sku_set
            )
        ]
        .groupby(
            "상품코드",
            as_index=False,
        )["현재고수량"]
        .sum()
        .rename(
            columns={
                "현재고수량": "기초재고"
            }
        )
    )

    i = (
        filtered[
            filtered["출입구분"] == "입고"
        ]
        .groupby(
            "상품코드",
            as_index=False,
        )["수량"]
        .sum()
        .rename(
            columns={
                "수량": "입고수량"
            }
        )
    )

    o = (
        filtered[
            filtered["출입구분"] == "출고"
        ]
        .groupby(
            "상품코드",
            as_index=False,
        )["수량"]
        .sum()
        .rename(
            columns={
                "수량": "출고수량"
            }
        )
    )

    d = (
        filtered[
            filtered["출입구분"] == "불량"
        ]
        .groupby(
            "상품코드",
            as_index=False,
        )["수량"]
        .sum()
        .rename(
            columns={
                "수량": "불량수량"
            }
        )
    )

    info = (
        filtered
        .groupby(
            "상품코드",
            as_index=False,
        )
        .agg(
            상품명=(
                "상품명",
                "first",
            ),
            공장=(
                "공장",
                "first",
            ),
            카테고리=(
                "카테고리",
                "first",
            ),
        )
    )

    result = sku.merge(
        info,
        on="상품코드",
        how="left",
    )

    for part in [
        b,
        i,
        o,
        d,
    ]:

        result = result.merge(
            part,
            on="상품코드",
            how="left",
        )

    for c in [
        "기초재고",
        "입고수량",
        "출고수량",
        "불량수량",
    ]:

        if c not in result.columns:
            result[c] = 0.0

        result[c] = (
            result[c]
            .fillna(0)
        )

    result["실제검수수량"] = (
        result["출고수량"]
        + result["불량수량"]
    )

    result["불량률"] = (
        result["불량수량"]
        .div(
            result["실제검수수량"]
            .replace(0, pd.NA)
        )
        * 100
    ).fillna(0)

    result["기말이론재고"] = (
        result["기초재고"]
        + result["입고수량"]
        - result["출고수량"]
        - result["불량수량"]
    )

    result["재고흐름"] = (
        result["기말이론재고"]
        .apply(
            lambda x:
            "점검필요(음수)"
            if x < 0
            else "정상"
        )
    )

    result["검수여부"] = (
        result["실제검수수량"]
        .apply(
            lambda x:
            "검수 발생"
            if x > 0
            else "미검수"
        )
    )

    return result


def stock_flow_validation(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
) -> pd.DataFrame:

    result = build_sku_summary(
        filtered,
        base,
    )

    if result.empty:
        return result

    result = result.sort_values(
        [
            "재고흐름",
            "기말이론재고",
        ],
        ascending=[
            True,
            True,
        ],
        key=lambda s: (
            s
            if s.name != "재고흐름"
            else s.map(
                {
                    "점검필요(음수)": 0,
                    "정상": 1,
                }
            )
        ),
    )

    return (
        result
        .reset_index(drop=True)
        .copy()
    )


# ==============================================================
# 차트
# ==============================================================

def line_chart(
    x,
    y,
    name,
    title,
    y_title,
    suffix="",
    color_light="#1565C0",
    color_dark="#64B5F6",
):

    fig = go.Figure()

    color = (
        color_dark
        if get_theme_type() == "dark"
        else color_light
    )

    fig.add_trace(
        go.Scatter(
            x=x,
            y=y,
            name=name,
            mode="lines+markers+text",
            text=[
                (
                    f"{v:,.2f}{suffix}"
                    if suffix
                    else f"{v:,.0f}"
                )
                for v in y
            ],
            textposition="top center",
            line=dict(
                color=color,
                width=4,
            ),
            marker=dict(
                size=8,
            ),
        )
    )

    fig = apply_theme(
        fig,
        420,
    )

    fig.update_layout(
        title=title,
        xaxis_title="년월",
        yaxis_title=y_title,
        yaxis=dict(
            ticksuffix=suffix,
        ),
    )

    return fig


# ==============================================================
# 화면 시작
# ==============================================================

st.title(
    "📦 입고 / 출고 / 불량 / 상품마스터 대시보드"
)

st.caption(
    "출고수량을 맞추기 위해 실제 검수한 수량을 기준으로 불량률을 계산합니다. "
    "예: 출고 100 + 불량 3 = 실제 검수 103 → 불량률 2.91%"
)


# ==============================================================
# 데이터 파일
# ==============================================================

data_source = (
    DEFAULT_DATA_PATH
    if os.path.exists(DEFAULT_DATA_PATH)
    else None
)

uploaded = st.file_uploader(
    "raw_data.xlsx 업로드 (선택)",
    type=["xlsx"],
    help="업로드하면 업로드한 파일을 사용합니다.",
)

if uploaded is not None:

    data_source = None
    file_bytes = uploaded.getvalue()

else:

    if data_source is None:

        st.info(
            "data/raw_data.xlsx를 넣거나 "
            "위에서 엑셀 파일을 업로드해주세요."
        )

        st.stop()

    with open(
        data_source,
        "rb",
    ) as f:

        file_bytes = f.read()


# ==============================================================
# 데이터 로딩
# ==============================================================

try:

    (
        df,
        base,
        master,
        validation,
    ) = load_and_process(
        file_bytes
    )

except Exception as e:

    st.error(
        f"데이터를 읽는 중 오류가 발생했습니다: {e}"
    )

    st.stop()


# ==============================================================
# 원본 데이터 / 상품마스터 / 기초재고 검증
# ==============================================================

with st.expander(
    "🔎 원본 데이터 / 상품마스터 / 기초재고 검증",
    expanded=True,
):

    # ----------------------------------------------------------
    # 상단 핵심 숫자
    # ----------------------------------------------------------

    a, b, c, d, e, f = st.columns(6)

    a.metric(
        "기초재고",
        f"{validation['base_qty']:,.0f}",
    )

    b.metric(
        "총 입고",
        f"{validation['source_qty'].get('입고', 0):,.0f}",
    )

    c.metric(
        "총 출고",
        f"{validation['source_qty'].get('출고', 0):,.0f}",
    )

    d.metric(
        "총 불량",
        f"{validation['source_qty'].get('불량', 0):,.0f}",
    )

    e.metric(
        "마스터 SKU",
        f"{validation['master']['master_unique_skus']:,}",
    )

    base_dates = validation[
        "base_date_values"
    ]

    f.metric(
        "기초재고 기준일",
        ", ".join(base_dates)
        if base_dates
        else "없음",
    )

    st.markdown("### 📋 원본 행 수")

    # 기존 st.write(dict) 제거
    # → 이상한 JSON/SVG 형태 숫자 출력 방지

    row_check = pd.DataFrame(
        [
            {
                "데이터": "입고",
                "원본 행 수": validation[
                    "source_rows"
                ].get("입고", 0),
            },
            {
                "데이터": "출고",
                "원본 행 수": validation[
                    "source_rows"
                ].get("출고", 0),
            },
            {
                "데이터": "불량",
                "원본 행 수": validation[
                    "source_rows"
                ].get("불량", 0),
            },
            {
                "데이터": "상품마스터",
                "원본 행 수": validation[
                    "master"
                ]["master_rows"],
            },
        ]
    )

    row_check = (
        row_check
        .reset_index(drop=True)
    )

    st.dataframe(
        hide_dataframe_index(
            row_check.style.format(
                {
                    "원본 행 수": "{:,.0f}"
                }
            )
        ),
        use_container_width=True,
        hide_index=True,
    )

    # ----------------------------------------------------------
    # 음수 수량
    # ----------------------------------------------------------

    st.markdown("### ⚠️ 음수 수량 행")

    negative_check = pd.DataFrame(
        [
            {
                "데이터": "불량",
                "음수 수량 행": validation[
                    "negative_qty_rows"
                ].get("불량", 0),
            },
            {
                "데이터": "입고",
                "음수 수량 행": validation[
                    "negative_qty_rows"
                ].get("입고", 0),
            },
            {
                "데이터": "출고",
                "음수 수량 행": validation[
                    "negative_qty_rows"
                ].get("출고", 0),
            },
            {
                "데이터": "기초재고",
                "음수 수량 행": validation[
                    "base_negative_rows"
                ],
            },
        ]
    )

    negative_check = (
        negative_check
        .reset_index(drop=True)
    )

    st.dataframe(
        hide_dataframe_index(
            negative_check.style.format(
                {
                    "음수 수량 행": "{:,.0f}"
                }
            )
        ),
        use_container_width=True,
        hide_index=True,
    )

    # ----------------------------------------------------------
    # 날짜 오류
    # ----------------------------------------------------------

    st.markdown("### 📅 날짜 오류 행")

    date_check = pd.DataFrame(
        [
            {
                "데이터": "불량",
                "날짜 오류 행": validation[
                    "invalid_dates"
                ].get("불량", 0),
            },
            {
                "데이터": "입고",
                "날짜 오류 행": validation[
                    "invalid_dates"
                ].get("입고", 0),
            },
            {
                "데이터": "출고",
                "날짜 오류 행": validation[
                    "invalid_dates"
                ].get("출고", 0),
            },
            {
                "데이터": "기초재고",
                "날짜 오류 행": validation[
                    "invalid_dates"
                ].get("기초재고", 0),
            },
        ]
    )

    date_check = (
        date_check
        .reset_index(drop=True)
    )

    st.dataframe(
        hide_dataframe_index(
            date_check.style.format(
                {
                    "날짜 오류 행": "{:,.0f}"
                }
            )
        ),
        use_container_width=True,
        hide_index=True,
    )

    # ----------------------------------------------------------
    # 상품마스터 미등록 SKU
    # ----------------------------------------------------------

    st.markdown(
        "### 🧩 상품마스터 미등록 SKU 수"
    )

    missing_master_check = pd.DataFrame(
        [
            {
                "데이터": "불량",
                "미등록 SKU 수": len(
                    validation[
                        "missing_master"
                    ].get(
                        "불량",
                        [],
                    )
                ),
            },
            {
                "데이터": "입고",
                "미등록 SKU 수": len(
                    validation[
                        "missing_master"
                    ].get(
                        "입고",
                        [],
                    )
                ),
            },
            {
                "데이터": "출고",
                "미등록 SKU 수": len(
                    validation[
                        "missing_master"
                    ].get(
                        "출고",
                        [],
                    )
                ),
            },
            {
                "데이터": "기초재고",
                "미등록 SKU 수": len(
                    validation[
                        "missing_master"
                    ].get(
                        "기초재고",
                        [],
                    )
                ),
            },
        ]
    )

    missing_master_check = (
        missing_master_check
        .reset_index(drop=True)
    )

    st.dataframe(
        hide_dataframe_index(
            missing_master_check.style.format(
                {
                    "미등록 SKU 수": "{:,.0f}"
                }
            )
        ),
        use_container_width=True,
        hide_index=True,
    )

    # ----------------------------------------------------------
    # 상세 검증
    # ----------------------------------------------------------

    if validation["master"][
        "master_duplicate_codes"
    ]:

        st.warning(
            "상품마스터에 상품코드 중복이 있습니다. "
            "첫 번째 행을 기준정보로 사용했습니다."
        )

        duplicate_table = pd.DataFrame(
            {
                "중복 상품코드":
                validation["master"][
                    "master_duplicate_codes"
                ][:100]
            }
        ).reset_index(drop=True)

        st.dataframe(
            hide_dataframe_index(
                duplicate_table.style
            ),
            use_container_width=True,
            hide_index=True,
        )

    missing_master_total = {
        k: len(v)
        for k, v in validation[
            "missing_master"
        ].items()
    }

    if any(
        v > 0
        for v in missing_master_total.values()
    ):

        st.warning(
            "상품마스터에 없는 상품코드는 "
            "공장/카테고리를 '미상'으로 처리했습니다. "
            "정확한 집계를 위해 상품마스터에 해당 SKU를 등록하는 것을 권장합니다."
        )

        missing_detail_rows = []

        for source_name, codes in validation[
            "missing_master"
        ].items():

            for code in codes:

                missing_detail_rows.append(
                    {
                        "데이터": source_name,
                        "미등록 상품코드": code,
                    }
                )

        if missing_detail_rows:

            missing_detail = pd.DataFrame(
                missing_detail_rows
            ).drop_duplicates()

            missing_detail = (
                missing_detail
                .reset_index(drop=True)
            )

            st.dataframe(
                hide_dataframe_index(
                    missing_detail.style
                ),
                use_container_width=True,
                height=250,
                hide_index=True,
            )

    else:

        st.success(
            "입고/출고/불량/기초재고의 상품코드가 "
            "모두 상품마스터와 연결되었습니다."
        )

    # ----------------------------------------------------------
    # 기초재고 기준일
    # ----------------------------------------------------------

    if (
        len(base_dates) == 1
        and base_dates[0] == "2026-01-01"
    ):

        st.success(
            "기초재고 기준일은 2026-01-01로 확인되었습니다."
        )

    else:

        st.warning(
            "기초재고 기준일이 2026-01-01 하나로 구성되어 있지 않습니다. "
            "현재 파일의 실제 기준일을 확인해주세요."
        )


# ==============================================================
# 필터
# ==============================================================

st.subheader("🔎 필터")

filter1, filter2, filter3, filter4 = st.columns(4)

factory_options = [
    x
    for x in FACTORIES
    if x in set(df["공장"])
]

category_options = [
    x
    for x in CATEGORIES
    if x in set(df["카테고리"])
]

defect_type_options = [
    x
    for x in DEFECT_TYPES_ORDER
    if x in set(df["불량타입"])
]

with filter1:

    selected_factories = st.multiselect(
        "🏭 공장",
        options=factory_options,
        default=factory_options,
    )

with filter2:

    selected_categories = st.multiselect(
        "👓 카테고리",
        options=category_options,
        default=category_options,
        help="G=안경 / S=선글라스",
    )

with filter3:

    selected_defect_types = st.multiselect(
        "🏷️ 불량타입",
        options=defect_type_options,
        default=defect_type_options,
    )

with filter4:

    selected_kinds = st.multiselect(
        "📂 데이터 종류",
        options=DATA_TYPES_ORDER,
        default=DATA_TYPES_ORDER,
    )


filter5, filter6, filter7 = st.columns(
    [1.4, 1.4, 1]
)

with filter5:

    product_name_query = st.text_input(
        "상품명 검색",
        placeholder="상품명",
    )

with filter6:

    product_code_query = st.text_input(
        "상품코드 검색",
        placeholder="상품코드",
    )

with filter7:

    min_date = (
        df["작업일"]
        .min()
        .date()
    )

    max_date = (
        df["작업일"]
        .max()
        .date()
    )

    date_range = st.date_input(
        "거래일 범위",
        value=(
            min_date,
            max_date,
        ),
        min_value=min_date,
        max_value=max_date,
    )


if not selected_factories:

    st.warning(
        "공장을 하나 이상 선택해주세요."
    )

    st.stop()


if not selected_categories:

    st.warning(
        "카테고리를 하나 이상 선택해주세요."
    )

    st.stop()


if not selected_kinds:

    st.warning(
        "데이터 종류를 하나 이상 선택해주세요."
    )

    st.stop()


if (
    not selected_defect_types
    and "불량" in selected_kinds
):

    st.warning(
        "불량 데이터를 보려면 "
        "불량타입을 하나 이상 선택해주세요."
    )

    st.stop()


if (
    isinstance(
        date_range,
        (tuple, list),
    )
    and len(date_range) == 2
):

    start_date, end_date = date_range

else:

    start_date = min_date
    end_date = max_date


filtered = filter_data(
    df,
    selected_factories,
    selected_categories,
    selected_defect_types,
    selected_kinds,
    start_date,
    end_date,
    product_name_query,
    product_code_query,
)


if filtered.empty:

    st.warning(
        "선택한 조건에 해당하는 거래 데이터가 없습니다."
    )

    st.stop()


# ==============================================================
# KPI
# ==============================================================

total_in = filtered.loc[
    filtered["출입구분"] == "입고",
    "수량",
].sum()

total_out = filtered.loc[
    filtered["출입구분"] == "출고",
    "수량",
].sum()

total_defect = filtered.loc[
    filtered["출입구분"] == "불량",
    "수량",
].sum()


base_filtered = filtered_base_stock(
    base,
    selected_factories,
    selected_categories,
    product_name_query,
    product_code_query,
)

total_base = base_filtered[
    "현재고수량"
].sum()

available_after_in = (
    total_base
    + total_in
)

ending_theoretical = (
    available_after_in
    - total_out
    - total_defect
)

# 실제 검수 = 정상 출고 + 불량
inspection_total = (
    total_out
    + total_defect
)

overall_defect_rate = safe_rate(
    total_defect,
    inspection_total,
)


k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

k1.metric(
    "선택 기초재고",
    f"{total_base:,.0f}",
)

k2.metric(
    "선택 입고",
    f"{total_in:,.0f}",
)

k3.metric(
    "선택 출고",
    f"{total_out:,.0f}",
)

k4.metric(
    "선택 불량",
    f"{total_defect:,.0f}",
)

k5.metric(
    "실제 검수수량",
    f"{inspection_total:,.0f}",
)

k6.metric(
    "검수 기준 불량률",
    f"{overall_defect_rate:.2f}%",
)

k7.metric(
    "이론 잔여재고",
    f"{ending_theoretical:,.0f}",
)


if ending_theoretical < 0:

    st.warning(
        "선택 조건에서 "
        "기초재고 + 입고 - 출고 - 불량이 음수입니다. "
        "출고/입고/불량 데이터 또는 기초재고를 확인해주세요."
    )


if inspection_total > 0:

    st.info(
        f"검수 기준: 실제 검수수량 = "
        f"출고 {total_out:,.0f} + "
        f"불량 {total_defect:,.0f} = "
        f"{inspection_total:,.0f} / "
        f"불량률 = "
        f"{total_defect:,.0f} ÷ "
        f"{inspection_total:,.0f} = "
        f"{overall_defect_rate:.2f}%"
    )

else:

    st.info(
        "현재 선택 조건에는 실제 검수수량 "
        "(출고+불량)이 없습니다."
    )


st.divider()


# ==============================================================
# 월별 재고 흐름
# ==============================================================

monthly = monthly_inventory_summary(
    filtered,
    base,
    selected_factories,
    selected_categories,
    product_name_query,
    product_code_query,
)

st.subheader(
    "📊 월별 재고 흐름 / 실제 검수 불량률"
)

fig_month = go.Figure()

dark = (
    get_theme_type() == "dark"
)


fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["입고"],
        name="입고",
        text=monthly["입고"].map(
            lambda x:
            f"{x:,.0f}"
        ),
        textposition="outside",
        marker_color=(
            "#42A5F5"
            if dark
            else "#1565C0"
        ),
        opacity=0.78,
    )
)


fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["출고"],
        name="출고",
        text=monthly["출고"].map(
            lambda x:
            f"{x:,.0f}"
        ),
        textposition="outside",
        marker_color=(
            "#66BB6A"
            if dark
            else "#2E7D32"
        ),
        opacity=0.82,
    )
)


fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["불량"],
        name="불량",
        text=monthly["불량"].map(
            lambda x:
            f"{x:,.0f}"
        ),
        textposition="outside",
        marker_color=(
            "#EF5350"
            if dark
            else "#C62828"
        ),
        opacity=0.86,
    )
)


fig_month.add_trace(
    go.Scatter(
        x=monthly["년월"],
        y=monthly["월불량률"],
        name="검수 기준 월불량률",
        mode="lines+markers+text",
        text=monthly["월불량률"].map(
            lambda x:
            f"{x:.2f}%"
        ),
        textposition="top center",
        line=dict(
            color=(
                "#FFD54F"
                if dark
                else "#E65100"
            ),
            width=3,
        ),
        marker=dict(
            size=8,
        ),
        yaxis="y2",
    )
)


fig_month = apply_theme(
    fig_month,
    540,
)

fig_month.update_layout(
    barmode="group",
    xaxis_title="년월",
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
        showgrid=False,
    ),
)

st.plotly_chart(
    fig_month,
    use_container_width=True,
)


st.caption(
    "핵심 기준: 월불량률 = "
    "당월 불량 ÷ (당월 출고 + 당월 불량). "
    "출고 100개를 확보하는 과정에서 불량 3개가 발생했다면 "
    "실제 검수는 103개입니다."
)


# ==============================================================
# 월별 누계
# ==============================================================

st.subheader(
    "📈 월별 누계"
)

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "누계 입고",
        "누계 출고",
        "누계 이론재고",
        "누계 불량률",
    ]
)


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
            monthly["누적출고"],
            "누계 출고",
            "누계 출고 수량",
            "수량",
            color_light="#2E7D32",
            color_dark="#81C784",
        ),
        use_container_width=True,
    )


with tab3:

    st.plotly_chart(
        line_chart(
            monthly["년월"],
            monthly["월말이론재고"],
            "월말 이론재고",
            "기초재고 + 누계입고 - 누계출고 - 누계불량",
            "이론재고",
            color_light="#6A1B9A",
            color_dark="#CE93D8",
        ),
        use_container_width=True,
    )


with tab4:

    st.plotly_chart(
        line_chart(
            monthly["년월"],
            monthly["누계불량률"],
            "누계 검수 기준 불량률",
            "누계 검수 기준 불량률",
            "불량률",
            suffix="%",
            color_light="#E65100",
            color_dark="#FFD54F",
        ),
        use_container_width=True,
    )


# ==============================================================
# 월별 계산 검증표
# ==============================================================

with st.expander(
    "📋 월별 계산 검증표",
    expanded=False,
):

    show_month = monthly[
        [
            "년월",
            "기초재고",
            "월초이론재고",
            "입고",
            "출고",
            "불량",
            "당월입고후재고",
            "월말이론재고",
            "실제검수수량",
            "월불량률",
            "누적입고",
            "누적출고",
            "누적불량",
            "누계검수수량",
            "누계불량률",
        ]
    ].copy()

    # 인덱스 명시적으로 초기화
    show_month = (
        show_month
        .reset_index(drop=True)
    )

    month_styler = show_month.style.format(
        {
            "기초재고": "{:,.0f}",
            "월초이론재고": "{:,.0f}",
            "입고": "{:,.0f}",
            "출고": "{:,.0f}",
            "불량": "{:,.0f}",
            "당월입고후재고": "{:,.0f}",
            "월말이론재고": "{:,.0f}",
            "실제검수수량": "{:,.0f}",
            "월불량률": "{:.2f}%",
            "누적입고": "{:,.0f}",
            "누적출고": "{:,.0f}",
            "누적불량": "{:,.0f}",
            "누계검수수량": "{:,.0f}",
            "누계불량률": "{:.2f}%",
        }
    )

    st.dataframe(
        hide_dataframe_index(
            month_styler
        ),
        use_container_width=True,
        hide_index=True,
    )


# ==============================================================
# 공장별
# ==============================================================

st.divider()

st.subheader(
    "🏭 공장별 / 카테고리별 입고 · 출고 · 불량 · 검수"
)

factory_summary = (
    filtered
    .groupby(
        ["공장", "출입구분"]
    )["수량"]
    .sum()
    .unstack(fill_value=0)
)

for col in DATA_TYPES_ORDER:

    if col not in factory_summary.columns:
        factory_summary[col] = 0

factory_summary = factory_summary[
    DATA_TYPES_ORDER
]


factory_summary["실제검수수량"] = (
    factory_summary["출고"]
    + factory_summary["불량"]
)

factory_summary["불량률"] = (
    factory_summary["불량"]
    .div(
        factory_summary["실제검수수량"]
        .replace(0, pd.NA)
    )
    * 100
).fillna(0)


factory_base = (
    base_filtered
    .groupby("공장")["현재고수량"]
    .sum()
    .reindex(
        factory_summary.index
    )
    .fillna(0)
)

factory_summary["기초재고"] = (
    factory_base
)

factory_summary["이론잔여재고"] = (
    factory_summary["기초재고"]
    + factory_summary["입고"]
    - factory_summary["출고"]
    - factory_summary["불량"]
)


factory_colors = (
    FACTORY_COLORS_DARK
    if dark
    else FACTORY_COLORS_LIGHT
)


fig_factory = go.Figure()


fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["입고"],
        name="입고",
        marker_color=[
            factory_colors.get(
                x,
                factory_colors["미상"],
            )
            for x in factory_summary.index
        ],
    )
)


fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["출고"],
        name="출고",
        marker_color=(
            "#66BB6A"
            if dark
            else "#2E7D32"
        ),
    )
)


fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["불량"],
        name="불량",
        marker_color=(
            "#EF5350"
            if dark
            else "#C62828"
        ),
    )
)


fig_factory.add_trace(
    go.Scatter(
        x=factory_summary.index,
        y=factory_summary["불량률"],
        name="검수 기준 불량률",
        mode="lines+markers+text",
        text=factory_summary["불량률"].map(
            lambda x:
            f"{x:.2f}%"
        ),
        textposition="top center",
        yaxis="y2",
        line=dict(
            color=(
                "#FFD54F"
                if dark
                else "#E65100"
            ),
            width=3,
        ),
    )
)


fig_factory = apply_theme(
    fig_factory,
    480,
)

fig_factory.update_layout(
    barmode="group",
    xaxis_title="공장",
    yaxis=dict(
        title="수량"
    ),
    yaxis2=dict(
        title="불량률 (%)",
        overlaying="y",
        side="right",
        rangemode="tozero",
        ticksuffix="%",
        showgrid=False,
    ),
)

st.plotly_chart(
    fig_factory,
    use_container_width=True,
)


factory_display = (
    factory_summary
    .reset_index()
    .copy()
)

factory_styler = factory_display.style.format(
    {
        "기초재고": "{:,.0f}",
        "입고": "{:,.0f}",
        "출고": "{:,.0f}",
        "불량": "{:,.0f}",
        "실제검수수량": "{:,.0f}",
        "불량률": "{:.2f}%",
        "이론잔여재고": "{:,.0f}",
    }
)

st.dataframe(
    hide_dataframe_index(
        factory_styler
    ),
    use_container_width=True,
    hide_index=True,
)


# ==============================================================
# 카테고리
# ==============================================================

category_summary = (
    filtered
    .groupby(
        ["카테고리", "출입구분"]
    )["수량"]
    .sum()
    .unstack(fill_value=0)
)

for col in DATA_TYPES_ORDER:

    if col not in category_summary.columns:
        category_summary[col] = 0

category_summary = category_summary[
    DATA_TYPES_ORDER
]


category_summary["실제검수수량"] = (
    category_summary["출고"]
    + category_summary["불량"]
)

category_summary["불량률"] = (
    category_summary["불량"]
    .div(
        category_summary["실제검수수량"]
        .replace(0, pd.NA)
    )
    * 100
).fillna(0)


category_base = (
    base_filtered
    .groupby("카테고리")["현재고수량"]
    .sum()
    .reindex(
        category_summary.index
    )
    .fillna(0)
)

category_summary["기초재고"] = (
    category_base
)

category_summary["이론잔여재고"] = (
    category_summary["기초재고"]
    + category_summary["입고"]
    - category_summary["출고"]
    - category_summary["불량"]
)


st.subheader(
    "👓 카테고리별 G / S"
)

category_display = (
    category_summary
    .reset_index()
    .copy()
)

category_styler = category_display.style.format(
    {
        "기초재고": "{:,.0f}",
        "입고": "{:,.0f}",
        "출고": "{:,.0f}",
        "불량": "{:,.0f}",
        "실제검수수량": "{:,.0f}",
        "불량률": "{:.2f}%",
        "이론잔여재고": "{:,.0f}",
    }
)

st.dataframe(
    hide_dataframe_index(
        category_styler
    ),
    use_container_width=True,
    hide_index=True,
)


# ==============================================================
# 불량타입
# ==============================================================

st.divider()

st.subheader(
    "🏷️ 불량타입별 월별 수량"
)

defect_only = filtered[
    filtered["출입구분"] == "불량"
].copy()


if defect_only.empty:

    st.info(
        "현재 조건에서는 불량 데이터가 없습니다."
    )

else:

    type_monthly = (
        defect_only
        .groupby(
            ["년월", "불량타입"]
        )["수량"]
        .sum()
        .unstack(fill_value=0)
        .sort_index()
    )

    for col in DEFECT_TYPES_ORDER:

        if col not in type_monthly.columns:
            type_monthly[col] = 0

    type_monthly = type_monthly[
        DEFECT_TYPES_ORDER
    ]


    tab_main, tab_minor, tab_table = st.tabs(
        [
            "주요 타입 · 테 / 렌즈",
            "군소 타입 · 전체 / 기타",
            "수량 검증표",
        ]
    )


    defect_colors = {
        "테": "#D32F2F",
        "렌즈": "#7B1FA2",
        "전체": "#6D4C41",
        "기타": "#616161",
    }


    if dark:

        defect_colors = {
            "테": "#EF5350",
            "렌즈": "#CE93D8",
            "전체": "#BCAAA4",
            "기타": "#B0BEC5",
        }


    def make_defect_group_chart(
        data,
        types,
        title,
    ):

        fig = go.Figure()

        for defect_type in types:

            fig.add_trace(
                go.Bar(
                    x=data.index,
                    y=data[defect_type],
                    name=defect_type,
                    marker_color=defect_colors[
                        defect_type
                    ],
                    text=data[
                        defect_type
                    ].map(
                        lambda x:
                        (
                            f"{x:,.0f}"
                            if x != 0
                            else ""
                        )
                    ),
                    textposition="outside",
                )
            )

        fig = apply_theme(
            fig,
            500,
        )

        fig.update_layout(
            title=title,
            barmode="group",
            xaxis_title="년월",
            yaxis_title="불량 수량",
        )

        return fig


    with tab_main:

        main_types = [
            x
            for x in [
                "테",
                "렌즈",
            ]
            if type_monthly[x]
            .abs()
            .sum() != 0
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

            st.info(
                "테/렌즈 데이터가 없습니다."
            )


    with tab_minor:

        minor_types = [
            x
            for x in [
                "전체",
                "기타",
            ]
            if type_monthly[x]
            .abs()
            .sum() != 0
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
                "전체/기타 데이터가 없습니다."
            )


    with tab_table:

        check_table = (
            type_monthly
            .copy()
        )

        check_table["불량 합계"] = (
            check_table.sum(axis=1)
        )

        check_table = (
            check_table
            .reset_index()
            .copy()
        )

        check_styler = check_table.style.format(
            {
                "테": "{:,.0f}",
                "렌즈": "{:,.0f}",
                "전체": "{:,.0f}",
                "기타": "{:,.0f}",
                "불량 합계": "{:,.0f}",
            }
        )

        st.dataframe(
            hide_dataframe_index(
                check_styler
            ),
            use_container_width=True,
            hide_index=True,
        )


# ==============================================================
# SKU별 상세 위험 / 검수
# ==============================================================

st.divider()

st.subheader(
    "🚨 SKU별 검수 / 불량 / 재고 현황"
)

s1, s2, s3 = st.columns(
    [1, 1, 1]
)


with s1:

    min_inspection_qty = st.number_input(
        "최소 실제 검수수량",
        min_value=0,
        value=10,
        step=10,
        help=(
            "실제 검수수량(출고+불량)이 "
            "너무 작은 SKU를 제외할 수 있습니다."
        ),
    )


with s2:

    min_available_qty = st.number_input(
        "최소 기초+입고량",
        min_value=0,
        value=0,
        step=10,
    )


with s3:

    top_n = st.number_input(
        "표시 SKU 수",
        min_value=5,
        max_value=200,
        value=50,
        step=5,
    )


sku_table = build_sku_summary(
    filtered,
    base,
)


if sku_table.empty:

    st.info(
        "현재 조건에서 SKU 데이터가 없습니다."
    )

else:

    sku_table = sku_table[
        (
            sku_table["실제검수수량"]
            >= min_inspection_qty
        )
        & (
            sku_table["기초재고"]
            + sku_table["입고수량"]
            >= min_available_qty
        )
    ].copy()


    sku_table = sku_table.sort_values(
        [
            "재고흐름",
            "불량률",
            "불량수량",
            "실제검수수량",
        ],
        ascending=[
            True,
            False,
            False,
            False,
        ],
        key=lambda s: (
            s
            if s.name != "재고흐름"
            else s.map(
                {
                    "점검필요(음수)": 0,
                    "정상": 1,
                }
            )
        ),
    )


    sku_table = (
        sku_table
        .head(int(top_n))
        .reset_index(drop=True)
        .copy()
    )


    if sku_table.empty:

        st.info(
            "현재 조건에서 표시할 SKU가 없습니다."
        )

    else:

        # 화면용 순위
        sku_table.insert(
            0,
            "순위",
            range(
                1,
                len(sku_table) + 1,
            ),
        )

        sku_display = sku_table[
            [
                "순위",
                "상품코드",
                "상품명",
                "공장",
                "카테고리",
                "기초재고",
                "입고수량",
                "출고수량",
                "불량수량",
                "실제검수수량",
                "불량률",
                "기말이론재고",
                "검수여부",
                "재고흐름",
            ]
        ].copy()


        sku_styler = sku_display.style.format(
            {
                "기초재고": "{:,.0f}",
                "입고수량": "{:,.0f}",
                "출고수량": "{:,.0f}",
                "불량수량": "{:,.0f}",
                "실제검수수량": "{:,.0f}",
                "불량률": "{:.2f}%",
                "기말이론재고": "{:,.0f}",
            }
        )


        st.dataframe(
            hide_dataframe_index(
                sku_styler
            ),
            use_container_width=True,
            hide_index=True,
        )


        st.caption(
            "불량률 = 불량수량 ÷ 실제 검수수량. "
            "실제 검수수량 = 출고수량 + 불량수량."
        )


# ==============================================================
# 재고 흐름 이상 SKU
# ==============================================================

st.divider()

st.subheader(
    "⚠️ 재고 흐름 점검"
)

flow_check = stock_flow_validation(
    filtered,
    base,
)


if flow_check.empty:

    st.info(
        "검증할 SKU가 없습니다."
    )

else:

    issue = flow_check[
        flow_check["재고흐름"]
        == "점검필요(음수)"
    ].copy()


    f1, f2, f3 = st.columns(3)


    f1.metric(
        "검증 SKU",
        f"{len(flow_check):,}",
    )


    f2.metric(
        "음수 이론재고 SKU",
        f"{len(issue):,}",
    )


    f3.metric(
        "음수 이론재고 합계",
        (
            f"{issue['기말이론재고'].sum():,.0f}"
            if not issue.empty
            else "0"
        ),
    )


    if issue.empty:

        st.success(
            "현재 선택 조건에서는 "
            "기초재고 + 입고 - 출고 - 불량이 "
            "음수가 되는 SKU가 없습니다."
        )

    else:

        st.warning(
            "아래 SKU는 기초재고 + 입고로 확보된 수량보다 "
            "출고+불량이 많습니다. 원본 거래를 확인해주세요."
        )


        issue_display = issue.head(
            200
        )[
            [
                "상품코드",
                "상품명",
                "공장",
                "카테고리",
                "기초재고",
                "입고수량",
                "출고수량",
                "불량수량",
                "실제검수수량",
                "기말이론재고",
            ]
        ].copy()


        issue_styler = issue_display.style.format(
            {
                "기초재고": "{:,.0f}",
                "입고수량": "{:,.0f}",
                "출고수량": "{:,.0f}",
                "불량수량": "{:,.0f}",
                "실제검수수량": "{:,.0f}",
                "기말이론재고": "{:,.0f}",
            }
        )


        st.dataframe(
            hide_dataframe_index(
                issue_styler
            ),
            use_container_width=True,
            hide_index=True,
        )


# ==============================================================
# 출고 기준 실제 검수 분석
# ==============================================================

st.divider()

st.subheader(
    "🔍 출고 기준 실제 검수 분석"
)

out_defect = filtered[
    filtered["출입구분"].isin(
        [
            "출고",
            "불량",
        ]
    )
].copy()


if out_defect.empty:

    st.info(
        "현재 조건에서는 출고/불량 검수 데이터가 없습니다."
    )

else:

    inspection_by_sku = (
        out_defect
        .groupby(
            [
                "상품코드",
                "상품명",
                "공장",
                "카테고리",
                "출입구분",
            ]
        )["수량"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )


    for col in [
        "출고",
        "불량",
    ]:

        if col not in inspection_by_sku.columns:
            inspection_by_sku[col] = 0


    inspection_by_sku["실제검수수량"] = (
        inspection_by_sku["출고"]
        + inspection_by_sku["불량"]
    )


    inspection_by_sku["불량률"] = (
        inspection_by_sku["불량"]
        .div(
            inspection_by_sku["실제검수수량"]
            .replace(0, pd.NA)
        )
        * 100
    ).fillna(0)


    inspection_by_sku = (
        inspection_by_sku
        .sort_values(
            [
                "불량률",
                "불량",
            ],
            ascending=[
                False,
                False,
            ],
        )
        .reset_index(drop=True)
        .copy()
    )


    inspection_display = inspection_by_sku[
        [
            "상품코드",
            "상품명",
            "공장",
            "카테고리",
            "출고",
            "불량",
            "실제검수수량",
            "불량률",
        ]
    ].copy()


    inspection_styler = inspection_display.style.format(
        {
            "출고": "{:,.0f}",
            "불량": "{:,.0f}",
            "실제검수수량": "{:,.0f}",
            "불량률": "{:.2f}%",
        }
    )


    st.dataframe(
        hide_dataframe_index(
            inspection_styler
        ),
        use_container_width=True,
        height=450,
        hide_index=True,
    )


# ==============================================================
# 상세 원본
# ==============================================================

st.divider()

st.subheader(
    "📄 필터링된 원본 데이터"
)

display_cols = [
    c
    for c in [
        "작업일",
        "출입구분",
        "공장",
        "카테고리",
        "불량타입",
        "상품코드",
        "상품명",
        "수량",
        "전표제목",
        "공급처 상품명",
        "현재고",
        "작업자",
        "전표번호",
    ]
    if c in filtered.columns
]


filtered_display = (
    filtered[
        display_cols
    ]
    .sort_values(
        [
            "작업일",
            "출입구분",
            "상품코드",
        ],
        ascending=[
            True,
            True,
            True,
        ],
    )
    .reset_index(drop=True)
    .copy()
)


# 화면용 원본 데이터도 인덱스 숨김
filtered_styler = (
    filtered_display
    .style
)


st.dataframe(
    hide_dataframe_index(
        filtered_styler
    ),
    use_container_width=True,
    height=500,
    hide_index=True,
)


# ==============================================================
# CSV
# ==============================================================

csv = filtered_display.to_csv(
    index=False,
    encoding="utf-8-sig",
)


st.download_button(
    "⬇️ 현재 필터 데이터 CSV 다운로드",
    data=csv,
    file_name="입고_출고_불량_필터데이터.csv",
    mime="text/csv",
)
