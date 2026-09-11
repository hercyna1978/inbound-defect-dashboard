# -*- coding: utf-8 -*-

"""
입고 / 출고 / 불량 / 기초재고 / 상품마스터 대시보드
=====================================================

핵심 집계 원칙
1. 입고
   - 입고 시트의 '수량'을 그대로 사용
   - 날짜는 '작업일' 사용

2. 출고
   - 출고 시트의 '출고수량'을 그대로 사용
   - '작업일'이 있으면 작업일 사용
   - 작업일이 없을 경우에만 '출고일' 사용

3. 불량
   - 불량 시트의 '수량'을 그대로 사용
   - 날짜는 '작업일' 사용

4. 상품마스터
   - 상품코드 기준으로 1:1 연결
   - 중복 상품코드는 첫 번째 행만 사용
   - 마스터 연결 때문에 거래 수량이 증가하지 않도록 처리

5. 실제 검수
   - 실제검수수량 = 출고 + 불량

6. 검수 기준 불량률
   - 불량률 = 불량 / (출고 + 불량)
   - 아직 검수하지 않은 재고는 분모에 포함하지 않음

7. 기초재고
   - 2026-01-01 기초재고 사용

8. 월말 이론재고
   - 월말이론재고 =
     기초재고 + 누계입고 - 누계출고 - 누계불량
   - '누계 이론재고'라는 표현은 사용하지 않음
"""


import io
import os
from typing import List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# Streamlit 기본 설정
# ============================================================

st.set_page_config(
    page_title="입고 / 출고 / 불량 대시보드",
    page_icon="📦",
    layout="wide",
)


# ============================================================
# 기본 설정
# ============================================================

DEFAULT_DATA_PATH = os.path.join(
    os.path.dirname(__file__),
    "data",
    "raw_data.xlsx",
)

BASE_DATE = pd.Timestamp("2026-01-01")


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


# ============================================================
# 필수 컬럼
# ============================================================

REQUIRED_INBOUND_COLUMNS = [
    "작업일",
    "수량",
    "상품코드",
    "상품명",
]


REQUIRED_DEFECT_COLUMNS = [
    "작업일",
    "수량",
    "상품코드",
    "상품명",
]


REQUIRED_OUTBOUND_COLUMNS = [
    "상품코드",
    "상품명",
    "출고수량",
]


REQUIRED_BASE_COLUMNS = [
    "기준일",
    "상품코드",
    "상품명",
    "현재고수량",
]


REQUIRED_MASTER_COLUMNS = [
    "상품코드",
    "상품명",
    "공장명",
    "카테고리",
]


# ============================================================
# 테마
# ============================================================

def get_theme_type() -> str:

    try:
        theme = getattr(
            st.context,
            "theme",
            None,
        )

        theme_type = getattr(
            theme,
            "type",
            None,
        )

        if theme_type in (
            "dark",
            "light",
        ):
            return theme_type

    except Exception:
        pass

    try:

        base = st.get_option(
            "theme.base"
        )

        if base in (
            "dark",
            "light",
        ):
            return base

    except Exception:
        pass

    return "light"


# ============================================================
# 차트 테마
# ============================================================

def apply_theme(
    fig: go.Figure,
    height: int = 450,
):

    dark = (
        get_theme_type()
        == "dark"
    )

    template = (
        "plotly_dark"
        if dark
        else "plotly_white"
    )

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
        font=dict(
            size=13
        ),
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


# ============================================================
# 기본 분류 함수
# ============================================================

def classify_factory(value) -> str:

    if pd.isna(value):
        return "미상"

    text = str(value).strip().upper()

    if not text:
        return "미상"

    if (
        "C2-S" in text
        or "C2S" in text
    ):
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

    if not text:
        return "기타"

    if "전체" in text:
        return "전체"

    if "렌즈" in text:
        return "렌즈"

    if "테" in text:
        return "테"

    return "기타"


# ============================================================
# 상품코드 정규화
# ============================================================

def normalize_code(
    series: pd.Series,
) -> pd.Series:

    s = (
        series
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # Excel 숫자형으로 읽힌 12345.0 제거
    s = s.str.replace(
        r"\.0$",
        "",
        regex=True,
    )

    return s


# ============================================================
# 숫자 변환
# ============================================================

def numeric_series(
    series: pd.Series,
) -> pd.Series:

    return pd.to_numeric(
        series,
        errors="coerce",
    ).fillna(0)


# ============================================================
# 상품마스터 준비
# ============================================================

def prepare_master(
    master_raw: pd.DataFrame,
) -> pd.DataFrame:

    missing = [
        c
        for c in REQUIRED_MASTER_COLUMNS
        if c not in master_raw.columns
    ]

    if missing:

        raise ValueError(
            "'상품마스터' 시트에 다음 컬럼이 없습니다: "
            f"{missing}"
        )

    master = master_raw.copy()

    master["상품코드"] = normalize_code(
        master["상품코드"]
    )

    master["상품명"] = (
        master["상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    master["공장명"] = (
        master["공장명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    master["카테고리"] = (
        master["카테고리"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    master = master[
        master["상품코드"].ne("")
    ].copy()

    # 상품코드 중복 제거
    master = master.drop_duplicates(
        subset=["상품코드"],
        keep="first",
    ).copy()

    master["공장"] = master[
        "공장명"
    ].apply(
        classify_factory
    )

    master["카테고리"] = master[
        "카테고리"
    ].apply(
        classify_category
    )

    return master


# ============================================================
# 상품마스터 연결
# ============================================================

def attach_master(
    df: pd.DataFrame,
    master: pd.DataFrame,
) -> pd.DataFrame:

    result = df.copy()

    result["상품코드"] = normalize_code(
        result["상품코드"]
    )

    # merge가 아닌 map 사용
    # → 상품마스터 중복 때문에 거래행이 복제되지 않음

    factory_map = (
        master
        .set_index("상품코드")["공장"]
        .to_dict()
    )

    category_map = (
        master
        .set_index("상품코드")["카테고리"]
        .to_dict()
    )

    master_name_map = (
        master
        .set_index("상품코드")["상품명"]
        .to_dict()
    )

    result["공장"] = (
        result["상품코드"]
        .map(factory_map)
        .fillna("미상")
    )

    result["카테고리"] = (
        result["상품코드"]
        .map(category_map)
        .fillna("미상")
    )

    master_names = (
        result["상품코드"]
        .map(master_name_map)
    )

    if "상품명" in result.columns:

        original_names = (
            result["상품명"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        result["상품명"] = (
            master_names
            .fillna("")
            .astype(str)
            .where(
                master_names
                .fillna("")
                .astype(str)
                .str.strip()
                .ne(""),
                original_names,
            )
        )

    return result


# ============================================================
# 원본 데이터 로딩
# ============================================================

@st.cache_data(
    show_spinner="엑셀 데이터를 불러오는 중입니다..."
)
def load_data(
    file_bytes: bytes,
) -> Tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    dict,
]:

    source = io.BytesIO(
        file_bytes
    )

    # ========================================================
    # 상품마스터
    # ========================================================

    try:

        master_raw = pd.read_excel(
            source,
            sheet_name="상품마스터",
        )

    except Exception as e:

        raise ValueError(
            f"'상품마스터' 시트를 읽을 수 없습니다: {e}"
        )

    master = prepare_master(
        master_raw
    )

    # ========================================================
    # 입고
    # ========================================================

    try:

        inbound_raw = pd.read_excel(
            source,
            sheet_name="입고",
        )

    except Exception as e:

        raise ValueError(
            f"'입고' 시트를 읽을 수 없습니다: {e}"
        )

    missing = [
        c
        for c in REQUIRED_INBOUND_COLUMNS
        if c not in inbound_raw.columns
    ]

    if missing:

        raise ValueError(
            f"'입고' 시트에 필수 컬럼이 없습니다: {missing}"
        )

    inbound = inbound_raw.copy()

    inbound["작업일"] = pd.to_datetime(
        inbound["작업일"],
        errors="coerce",
    )

    inbound["수량"] = numeric_series(
        inbound["수량"]
    )

    inbound["상품코드"] = normalize_code(
        inbound["상품코드"]
    )

    inbound["상품명"] = (
        inbound["상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    inbound["출입구분"] = "입고"

    if "전표제목" not in inbound.columns:
        inbound["전표제목"] = ""

    if "공급처 상품명" not in inbound.columns:
        inbound["공급처 상품명"] = ""

    inbound["전표제목"] = (
        inbound["전표제목"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    inbound["공급처 상품명"] = (
        inbound["공급처 상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    inbound["불량타입"] = "기타"

    # ========================================================
    # 불량
    # ========================================================

    try:

        defect_raw = pd.read_excel(
            source,
            sheet_name="불량",
        )

    except Exception as e:

        raise ValueError(
            f"'불량' 시트를 읽을 수 없습니다: {e}"
        )

    missing = [
        c
        for c in REQUIRED_DEFECT_COLUMNS
        if c not in defect_raw.columns
    ]

    if missing:

        raise ValueError(
            f"'불량' 시트에 필수 컬럼이 없습니다: {missing}"
        )

    defect = defect_raw.copy()

    defect["작업일"] = pd.to_datetime(
        defect["작업일"],
        errors="coerce",
    )

    defect["수량"] = numeric_series(
        defect["수량"]
    )

    defect["상품코드"] = normalize_code(
        defect["상품코드"]
    )

    defect["상품명"] = (
        defect["상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    defect["출입구분"] = "불량"

    if "전표제목" not in defect.columns:
        defect["전표제목"] = ""

    if "공급처 상품명" not in defect.columns:
        defect["공급처 상품명"] = ""

    defect["전표제목"] = (
        defect["전표제목"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    defect["공급처 상품명"] = (
        defect["공급처 상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    defect["불량타입"] = defect[
        "전표제목"
    ].apply(
        classify_defect_type
    )

    # ========================================================
    # 출고
    # ========================================================

    try:

        outbound_raw = pd.read_excel(
            source,
            sheet_name="출고",
        )

    except Exception as e:

        raise ValueError(
            f"'출고' 시트를 읽을 수 없습니다: {e}"
        )

    missing = [
        c
        for c in REQUIRED_OUTBOUND_COLUMNS
        if c not in outbound_raw.columns
    ]

    if missing:

        raise ValueError(
            f"'출고' 시트에 필수 컬럼이 없습니다: {missing}"
        )

    outbound = outbound_raw.copy()

    # --------------------------------------------------------
    # 출고 날짜
    #
    # 작업일이 있으면 작업일 사용
    # 작업일이 없으면 출고일 사용
    # --------------------------------------------------------

    if "작업일" in outbound.columns:

        outbound["작업일"] = pd.to_datetime(
            outbound["작업일"],
            errors="coerce",
        )

    else:

        if "출고일" not in outbound.columns:

            raise ValueError(
                "'출고' 시트에 '작업일' 또는 '출고일'이 필요합니다."
            )

        outbound["작업일"] = pd.to_datetime(
            outbound["출고일"],
            errors="coerce",
        )

    # 원본 출고수량 그대로 사용
    outbound["출고수량"] = numeric_series(
        outbound["출고수량"]
    )

    outbound["상품코드"] = normalize_code(
        outbound["상품코드"]
    )

    outbound["상품명"] = (
        outbound["상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    outbound["출입구분"] = "출고"

    outbound["수량"] = outbound[
        "출고수량"
    ]

    outbound["불량타입"] = "기타"

    if "전표제목" not in outbound.columns:
        outbound["전표제목"] = ""

    if "공급처 상품명" not in outbound.columns:
        outbound["공급처 상품명"] = ""

    outbound["전표제목"] = (
        outbound["전표제목"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    outbound["공급처 상품명"] = (
        outbound["공급처 상품명"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ========================================================
    # 원본 수량 먼저 계산
    # ========================================================

    raw_totals = {
        "입고": float(
            inbound["수량"].sum()
        ),
        "출고": float(
            outbound["출고수량"].sum()
        ),
        "불량": float(
            defect["수량"].sum()
        ),
    }

    raw_rows = {
        "입고": len(inbound),
        "출고": len(outbound),
        "불량": len(defect),
    }

    # ========================================================
    # 마스터 연결
    # ========================================================

    inbound = attach_master(
        inbound,
        master,
    )

    outbound = attach_master(
        outbound,
        master,
    )

    defect = attach_master(
        defect,
        master,
    )

    # ========================================================
    # 마스터 연결 후 수량 불변 확인
    # ========================================================

    processed_totals = {
        "입고": float(
            inbound["수량"].sum()
        ),
        "출고": float(
            outbound["수량"].sum()
        ),
        "불량": float(
            defect["수량"].sum()
        ),
    }

    for kind in DATA_TYPES_ORDER:

        if abs(
            raw_totals[kind]
            - processed_totals[kind]
        ) > 0.000001:

            raise ValueError(
                f"{kind} 원본 수량과 처리 후 수량이 "
                f"달라졌습니다. "
                f"원본={raw_totals[kind]:,.0f}, "
                f"처리후={processed_totals[kind]:,.0f}"
            )

    # ========================================================
    # 거래 데이터 통합
    # ========================================================

    transaction_frames = [
        inbound,
        outbound,
        defect,
    ]

    df = pd.concat(
        transaction_frames,
        ignore_index=True,
        sort=False,
    )

    df = df[
        df["작업일"].notna()
    ].copy()

    df["년월"] = (
        df["작업일"]
        .dt.to_period("M")
        .astype(str)
    )

    # ========================================================
    # 기초재고
    # ========================================================

    try:

        base_raw = pd.read_excel(
            source,
            sheet_name="기초재고",
        )

    except Exception as e:

        raise ValueError(
            f"'기초재고' 시트를 읽을 수 없습니다: {e}"
        )

    missing = [
        c
        for c in REQUIRED_BASE_COLUMNS
        if c not in base_raw.columns
    ]

    if missing:

        raise ValueError(
            f"'기초재고' 시트에 필수 컬럼이 없습니다: {missing}"
        )

    base = base_raw.copy()

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

    base["현재고수량"] = numeric_series(
        base["현재고수량"]
    )

    base = base[
        base["기준일"].notna()
    ].copy()

    base = attach_master(
        base,
        master,
    )

    # ========================================================
    # 기초재고 기준일
    # ========================================================

    exact_base = base[
        base["기준일"].eq(BASE_DATE)
    ].copy()

    if not exact_base.empty:

        base_selected = exact_base.copy()

    else:

        before_base = base[
            base["기준일"].le(BASE_DATE)
        ].copy()

        if not before_base.empty:

            latest_date = before_base[
                "기준일"
            ].max()

            base_selected = before_base[
                before_base["기준일"].eq(
                    latest_date
                )
            ].copy()

        else:

            base_selected = base.copy()

    # ========================================================
    # 내부 데이터 무결성 정보
    # ========================================================

    validation = {
        "raw_totals": raw_totals,
        "processed_totals": processed_totals,
        "raw_rows": raw_rows,
        "master_rows": len(master_raw),
        "master_unique_skus": master[
            "상품코드"
        ].nunique(),
        "base_qty": float(
            base_selected[
                "현재고수량"
            ].sum()
        ),
        "base_date": (
            base_selected["기준일"]
            .iloc[0]
            .strftime("%Y-%m-%d")
            if not base_selected.empty
            else ""
        ),
        "negative_rows": {
            "입고": int(
                (
                    inbound["수량"]
                    < 0
                ).sum()
            ),
            "출고": int(
                (
                    outbound["수량"]
                    < 0
                ).sum()
            ),
            "불량": int(
                (
                    defect["수량"]
                    < 0
                ).sum()
            ),
        },
    }

    return (
        df,
        base_selected,
        master,
        validation,
    )


# ============================================================
# 안전한 비율 계산
# ============================================================

def safe_rate(
    numerator,
    denominator,
) -> float:

    if denominator == 0:
        return 0.0

    return (
        float(numerator)
        / float(denominator)
        * 100
    )


# ============================================================
# 거래 필터
# ============================================================

def filter_transactions(
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
        df["공장"].isin(
            selected_factories
        )
        &
        df["카테고리"].isin(
            selected_categories
        )
        &
        df["출입구분"].isin(
            selected_kinds
        )
        &
        (
            df["작업일"].dt.date
            >= start_date
        )
        &
        (
            df["작업일"].dt.date
            <= end_date
        )
    )

    result = df.loc[
        mask
    ].copy()

    if not result.empty:

        defect_mask = (
            result["출입구분"]
            .eq("불량")
        )

        result = result[
            (~defect_mask)
            |
            result["불량타입"].isin(
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


# ============================================================
# 기초재고 필터
# ============================================================

def filter_base(
    base: pd.DataFrame,
    selected_factories: List[str],
    selected_categories: List[str],
    product_name_query: str,
    product_code_query: str,
) -> pd.DataFrame:

    result = base[
        base["공장"].isin(
            selected_factories
        )
        &
        base["카테고리"].isin(
            selected_categories
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


# ============================================================
# 월별 집계
# ============================================================

def make_monthly_summary(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
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

    # --------------------------------------------------------
    # 월별 수량
    # --------------------------------------------------------

    monthly = (
        filtered
        .groupby(
            [
                "년월",
                "출입구분",
            ],
            dropna=False,
        )["수량"]
        .sum()
        .unstack(
            fill_value=0
        )
        .sort_index()
    )

    for kind in DATA_TYPES_ORDER:

        if kind not in monthly.columns:
            monthly[kind] = 0.0

    monthly = monthly[
        DATA_TYPES_ORDER
    ].copy()

    # --------------------------------------------------------
    # 기초재고
    # --------------------------------------------------------

    base_qty = float(
        base[
            "현재고수량"
        ].sum()
    )

    monthly["기초재고"] = base_qty

    # --------------------------------------------------------
    # 누계
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # 이론재고
    # --------------------------------------------------------

    monthly["월초이론재고"] = (
        base_qty
        + monthly["누적입고"].shift(
            1,
            fill_value=0,
        )
        - monthly["누적출고"].shift(
            1,
            fill_value=0,
        )
        - monthly["누적불량"].shift(
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

    # --------------------------------------------------------
    # 실제 검수
    # --------------------------------------------------------

    monthly["실제검수수량"] = (
        monthly["출고"]
        + monthly["불량"]
    )

    monthly["월불량률"] = (
        monthly["불량"]
        .div(
            monthly[
                "실제검수수량"
            ].replace(
                0,
                pd.NA,
            )
        )
        * 100
    ).fillna(0)

    # --------------------------------------------------------
    # 누계 검수
    # --------------------------------------------------------

    monthly["누계검수수량"] = (
        monthly["누적출고"]
        + monthly["누적불량"]
    )

    monthly["누계불량률"] = (
        monthly["누적불량"]
        .div(
            monthly[
                "누계검수수량"
            ].replace(
                0,
                pd.NA,
            )
        )
        * 100
    ).fillna(0)

    monthly.index.name = "년월"

    result = monthly.reset_index()

    return result


# ============================================================
# SKU 요약
# ============================================================

def make_sku_summary(
    filtered: pd.DataFrame,
    base: pd.DataFrame,
) -> pd.DataFrame:

    if filtered.empty:
        return pd.DataFrame()

    sku_codes = sorted(
        set(
            filtered["상품코드"]
            .astype(str)
        )
        - {""}
    )

    if not sku_codes:
        return pd.DataFrame()

    result = pd.DataFrame(
        {
            "상품코드": sku_codes
        }
    )

    # --------------------------------------------------------
    # 상품 기본정보
    # --------------------------------------------------------

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

    result = result.merge(
        info,
        on="상품코드",
        how="left",
    )

    # --------------------------------------------------------
    # 기초재고
    # --------------------------------------------------------

    base_sku = (
        base[
            base["상품코드"]
            .isin(sku_codes)
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

    result = result.merge(
        base_sku,
        on="상품코드",
        how="left",
    )

    # --------------------------------------------------------
    # 입고
    # --------------------------------------------------------

    inbound = (
        filtered[
            filtered["출입구분"]
            .eq("입고")
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

    result = result.merge(
        inbound,
        on="상품코드",
        how="left",
    )

    # --------------------------------------------------------
    # 출고
    # --------------------------------------------------------

    outbound = (
        filtered[
            filtered["출입구분"]
            .eq("출고")
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

    result = result.merge(
        outbound,
        on="상품코드",
        how="left",
    )

    # --------------------------------------------------------
    # 불량
    # --------------------------------------------------------

    defect = (
        filtered[
            filtered["출입구분"]
            .eq("불량")
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

    result = result.merge(
        defect,
        on="상품코드",
        how="left",
    )

    for col in [
        "기초재고",
        "입고수량",
        "출고수량",
        "불량수량",
    ]:

        if col not in result.columns:
            result[col] = 0.0

        result[col] = (
            result[col]
            .fillna(0)
        )

    # --------------------------------------------------------
    # 실제 검수
    # --------------------------------------------------------

    result["실제검수수량"] = (
        result["출고수량"]
        + result["불량수량"]
    )

    # --------------------------------------------------------
    # 검수 기준 불량률
    # --------------------------------------------------------

    result["불량률"] = (
        result["불량수량"]
        .div(
            result[
                "실제검수수량"
            ].replace(
                0,
                pd.NA,
            )
        )
        * 100
    ).fillna(0)

    # --------------------------------------------------------
    # 이론재고
    # --------------------------------------------------------

    result["기말이론재고"] = (
        result["기초재고"]
        + result["입고수량"]
        - result["출고수량"]
        - result["불량수량"]
    )

    return result


# ============================================================
# 라인차트
# ============================================================

def make_line_chart(
    x,
    y,
    name,
    title,
    y_title,
    suffix="",
):

    dark = (
        get_theme_type()
        == "dark"
    )

    if dark:
        line_color = "#64B5F6"
    else:
        line_color = "#1565C0"

    fig = go.Figure()

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
                    else
                    f"{v:,.0f}"
                )
                for v in y
            ],
            textposition="top center",
            line=dict(
                color=line_color,
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
            separatethousands=True,
        ),
    )

    return fig


# ============================================================
# 데이터프레임 출력
# ============================================================

def show_table(
    df: pd.DataFrame,
    height=None,
):

    if df is None or df.empty:

        st.info(
            "표시할 데이터가 없습니다."
        )

        return

    # --------------------------------------------------------
    # 숫자 컬럼 표시 형식
    #
    # 실제 데이터는 숫자형 그대로 유지
    # 화면에서만 천 단위 콤마 표시
    # --------------------------------------------------------

    column_config = {}

    percent_keywords = [
        "불량률",
    ]

    for col in df.columns:

        if (
            pd.api.types.is_numeric_dtype(
                df[col]
            )
        ):

            if any(
                keyword in str(col)
                for keyword in percent_keywords
            ):

                column_config[col] = (
                    st.column_config.NumberColumn(
                        str(col),
                        format="%.2f%%",
                    )
                )

            else:

                column_config[col] = (
                    st.column_config.NumberColumn(
                        str(col),
                        format="%,.0f",
                    )
                )

    kwargs = {
        "use_container_width": True,
        "hide_index": True,
        "column_config": column_config,
    }

    if height is not None:
        kwargs["height"] = height

    st.dataframe(
        df,
        **kwargs,
    )


# ============================================================
# 화면
# ============================================================

st.title(
    "📦 입고 / 출고 / 불량 / 상품마스터 대시보드"
)

st.caption(
    "입고·출고·불량은 원본 거래 수량을 그대로 집계하며, "
    "검수 기준 불량률은 불량 ÷ (출고 + 불량)으로 계산합니다."
)


# ============================================================
# 파일
# ============================================================

data_source = (
    DEFAULT_DATA_PATH
    if os.path.exists(
        DEFAULT_DATA_PATH
    )
    else None
)


uploaded = st.file_uploader(
    "raw_data.xlsx 업로드",
    type=["xlsx"],
    help=(
        "업로드하면 업로드한 엑셀 파일을 사용합니다."
    ),
)


if uploaded is not None:

    file_bytes = (
        uploaded.getvalue()
    )

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


# ============================================================
# 데이터 처리
# ============================================================

try:

    (
        df,
        base,
        master,
        validation,
    ) = load_data(
        file_bytes
    )

except Exception as e:

    st.error(
        f"데이터를 읽는 중 오류가 발생했습니다: {e}"
    )

    st.stop()


# ============================================================
# 필터
# ============================================================

st.subheader(
    "🔎 필터"
)


filter1, filter2, filter3, filter4 = st.columns(4)


factory_options = [
    x
    for x in FACTORIES
    if x in set(
        df["공장"].dropna()
    )
]


category_options = [
    x
    for x in CATEGORIES
    if x in set(
        df["카테고리"].dropna()
    )
]


defect_type_options = [
    x
    for x in DEFECT_TYPES_ORDER
    if x in set(
        df["불량타입"].dropna()
    )
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
        help="G = 안경 / S = 선글라스",
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


# ============================================================
# 필터 예외 처리
# ============================================================

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

    start_date = date_range[0]
    end_date = date_range[1]

else:

    start_date = min_date
    end_date = max_date


# ============================================================
# 필터 적용
# ============================================================

filtered = filter_transactions(
    df=df,
    selected_factories=selected_factories,
    selected_categories=selected_categories,
    selected_defect_types=selected_defect_types,
    selected_kinds=selected_kinds,
    start_date=start_date,
    end_date=end_date,
    product_name_query=product_name_query,
    product_code_query=product_code_query,
)


if filtered.empty:

    st.warning(
        "선택한 조건에 해당하는 거래 데이터가 없습니다."
    )

    st.stop()


# ============================================================
# 기초재고 필터
# ============================================================

base_filtered = filter_base(
    base=base,
    selected_factories=selected_factories,
    selected_categories=selected_categories,
    product_name_query=product_name_query,
    product_code_query=product_code_query,
)


# ============================================================
# KPI
# ============================================================

total_in = float(
    filtered.loc[
        filtered["출입구분"].eq("입고"),
        "수량",
    ].sum()
)


total_out = float(
    filtered.loc[
        filtered["출입구분"].eq("출고"),
        "수량",
    ].sum()
)


total_defect = float(
    filtered.loc[
        filtered["출입구분"].eq("불량"),
        "수량",
    ].sum()
)


total_base = float(
    base_filtered[
        "현재고수량"
    ].sum()
)


inspection_total = (
    total_out
    + total_defect
)


overall_defect_rate = safe_rate(
    total_defect,
    inspection_total,
)


ending_theoretical = (
    total_base
    + total_in
    - total_out
    - total_defect
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


if inspection_total > 0:

    st.info(
        f"실제 검수수량 = "
        f"출고 {total_out:,.0f} + "
        f"불량 {total_defect:,.0f} = "
        f"{inspection_total:,.0f} / "
        f"검수 기준 불량률 = "
        f"{total_defect:,.0f} ÷ "
        f"{inspection_total:,.0f} = "
        f"{overall_defect_rate:.2f}%"
    )


st.divider()


# ============================================================
# 월별 집계
# ============================================================

monthly = make_monthly_summary(
    filtered=filtered,
    base=base_filtered,
)


# ============================================================
# 월별 재고 흐름
# ============================================================

st.subheader(
    "📊 월별 입고 · 출고 · 불량"
)


dark = (
    get_theme_type()
    == "dark"
)


fig_month = go.Figure()


fig_month.add_trace(
    go.Bar(
        x=monthly["년월"],
        y=monthly["입고"],
        name="입고",
        text=monthly["입고"].map(
            lambda x: f"{x:,.0f}"
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
            lambda x: f"{x:,.0f}"
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
            lambda x: f"{x:,.0f}"
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
        text=monthly[
            "월불량률"
        ].map(
            lambda x: f"{x:.2f}%"
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
            size=8
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
    "월별 수량은 각 원본 시트의 실제 수량을 그대로 합산합니다. "
    "검수 기준 월불량률 = 당월 불량 ÷ (당월 출고 + 당월 불량)"
)


# ============================================================
# 월별 계산표
# ============================================================

st.subheader(
    "📋 월별 수량 / 누계"
)


monthly_display = monthly[
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


# ------------------------------------------------------------
# 중요
#
# 숫자를 문자열로 변환하지 않습니다.
#
# show_table()에서
# Streamlit NumberColumn format을 이용해서
# 화면에만 천 단위 콤마를 표시합니다.
#
# 따라서
# 238695 → 238,695
# 1292387 → 1,292,387
#
# 이지만 내부 계산값은 계속 숫자형입니다.
# ------------------------------------------------------------

show_table(
    monthly_display
)


st.divider()


# ============================================================
# 월별 누계 차트
# ============================================================

st.subheader(
    "📈 월별 누계"
)


tab1, tab2, tab3, tab4 = st.tabs(
    [
        "누계 입고",
        "누계 출고",
        "월말 이론재고",
        "누계 불량률",
    ]
)


with tab1:

    st.plotly_chart(
        make_line_chart(
            monthly["년월"],
            monthly["누적입고"],
            "누계 입고",
            "누계 입고",
            "수량",
        ),
        use_container_width=True,
    )


with tab2:

    st.plotly_chart(
        make_line_chart(
            monthly["년월"],
            monthly["누적출고"],
            "누계 출고",
            "누계 출고",
            "수량",
        ),
        use_container_width=True,
    )


with tab3:

    st.plotly_chart(
        make_line_chart(
            monthly["년월"],
            monthly["월말이론재고"],
            "월말 이론재고",
            "월말 이론재고",
            "이론재고",
        ),
        use_container_width=True,
    )


with tab4:

    st.plotly_chart(
        make_line_chart(
            monthly["년월"],
            monthly["누계불량률"],
            "누계 검수 기준 불량률",
            "누계 검수 기준 불량률",
            "불량률",
            suffix="%",
        ),
        use_container_width=True,
    )


st.divider()


# ============================================================
# 공장별
# ============================================================

st.subheader(
    "🏭 공장별 입고 · 출고 · 불량"
)


factory_summary = (
    filtered
    .groupby(
        [
            "공장",
            "출입구분",
        ]
    )["수량"]
    .sum()
    .unstack(
        fill_value=0
    )
)


for kind in DATA_TYPES_ORDER:

    if kind not in factory_summary.columns:

        factory_summary[kind] = 0.0


factory_summary = factory_summary[
    DATA_TYPES_ORDER
].copy()


factory_summary[
    "실제검수수량"
] = (
    factory_summary["출고"]
    + factory_summary["불량"]
)


factory_summary[
    "불량률"
] = (
    factory_summary["불량"]
    .div(
        factory_summary[
            "실제검수수량"
        ].replace(
            0,
            pd.NA,
        )
    )
    * 100
).fillna(0)


factory_base = (
    base_filtered
    .groupby("공장")[
        "현재고수량"
    ]
    .sum()
    .reindex(
        factory_summary.index
    )
    .fillna(0)
)


factory_summary[
    "기초재고"
] = factory_base


factory_summary[
    "이론잔여재고"
] = (
    factory_summary["기초재고"]
    + factory_summary["입고"]
    - factory_summary["출고"]
    - factory_summary["불량"]
)


fig_factory = go.Figure()


fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["입고"],
        name="입고",
    )
)


fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["출고"],
        name="출고",
    )
)


fig_factory.add_trace(
    go.Bar(
        x=factory_summary.index,
        y=factory_summary["불량"],
        name="불량",
    )
)


fig_factory.add_trace(
    go.Scatter(
        x=factory_summary.index,
        y=factory_summary["불량률"],
        name="검수 기준 불량률",
        mode="lines+markers+text",
        text=factory_summary[
            "불량률"
        ].map(
            lambda x: f"{x:.2f}%"
        ),
        textposition="top center",
        yaxis="y2",
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
    fig_factory,
    use_container_width=True,
)


factory_display = (
    factory_summary
    .reset_index()
)


show_table(
    factory_display
)


# ============================================================
# 카테고리
# ============================================================

st.subheader(
    "👓 카테고리별 G / S"
)


category_summary = (
    filtered
    .groupby(
        [
            "카테고리",
            "출입구분",
        ]
    )["수량"]
    .sum()
    .unstack(
        fill_value=0
    )
)


for kind in DATA_TYPES_ORDER:

    if kind not in category_summary.columns:

        category_summary[kind] = 0.0


category_summary = category_summary[
    DATA_TYPES_ORDER
].copy()


category_summary[
    "실제검수수량"
] = (
    category_summary["출고"]
    + category_summary["불량"]
)


category_summary[
    "불량률"
] = (
    category_summary["불량"]
    .div(
        category_summary[
            "실제검수수량"
        ].replace(
            0,
            pd.NA,
        )
    )
    * 100
).fillna(0)


category_base = (
    base_filtered
    .groupby("카테고리")[
        "현재고수량"
    ]
    .sum()
    .reindex(
        category_summary.index
    )
    .fillna(0)
)


category_summary[
    "기초재고"
] = category_base


category_summary[
    "이론잔여재고"
] = (
    category_summary["기초재고"]
    + category_summary["입고"]
    - category_summary["출고"]
    - category_summary["불량"]
)


category_display = (
    category_summary
    .reset_index()
)


show_table(
    category_display
)


# ============================================================
# 불량타입별 월별
# ============================================================

st.divider()


st.subheader(
    "🏷️ 불량타입별 월별 수량"
)


defect_only = filtered[
    filtered["출입구분"].eq("불량")
].copy()


if defect_only.empty:

    st.info(
        "현재 조건에서는 불량 데이터가 없습니다."
    )

else:

    type_monthly = (
        defect_only
        .groupby(
            [
                "년월",
                "불량타입",
            ]
        )["수량"]
        .sum()
        .unstack(
            fill_value=0
        )
        .sort_index()
    )

    for defect_type in DEFECT_TYPES_ORDER:

        if (
            defect_type
            not in type_monthly.columns
        ):

            type_monthly[
                defect_type
            ] = 0.0

    type_monthly = type_monthly[
        DEFECT_TYPES_ORDER
    ].copy()

    tab_main, tab_minor, tab_table = st.tabs(
        [
            "주요 타입 · 테 / 렌즈",
            "군소 타입 · 전체 / 기타",
            "수량표",
        ]
    )

    def make_defect_chart(
        data,
        defect_types,
        title,
    ):

        fig = go.Figure()

        for defect_type in defect_types:

            fig.add_trace(
                go.Bar(
                    x=data.index,
                    y=data[
                        defect_type
                    ],
                    name=defect_type,
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
            if type_monthly[
                x
            ].abs().sum() != 0
        ]

        if main_types:

            st.plotly_chart(
                make_defect_chart(
                    type_monthly,
                    main_types,
                    "테 / 렌즈 월별 불량 수량",
                ),
                use_container_width=True,
            )

        else:

            st.info(
                "테 / 렌즈 데이터가 없습니다."
            )

    with tab_minor:

        minor_types = [
            x
            for x in [
                "전체",
                "기타",
            ]
            if type_monthly[
                x
            ].abs().sum() != 0
        ]

        if minor_types:

            st.plotly_chart(
                make_defect_chart(
                    type_monthly,
                    minor_types,
                    "전체 / 기타 월별 불량 수량",
                ),
                use_container_width=True,
            )

        else:

            st.info(
                "전체 / 기타 데이터가 없습니다."
            )

    with tab_table:

        defect_table = (
            type_monthly
            .reset_index()
        )

        defect_table[
            "불량 합계"
        ] = (
            defect_table[
                DEFECT_TYPES_ORDER
            ].sum(axis=1)
        )

        show_table(
            defect_table
        )


# ============================================================
# SKU별 현황
# ============================================================

st.divider()


st.subheader(
    "🚨 SKU별 검수 / 불량 / 재고 현황"
)


s1, s2, s3 = st.columns(3)


with s1:

    min_inspection_qty = st.number_input(
        "최소 실제 검수수량",
        min_value=0,
        value=10,
        step=10,
        help=(
            "실제 검수수량(출고+불량)이 "
            "너무 작은 SKU를 제외합니다."
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


sku_table = make_sku_summary(
    filtered=filtered,
    base=base_filtered,
)


if sku_table.empty:

    st.info(
        "현재 조건에서 SKU 데이터가 없습니다."
    )

else:

    sku_table = sku_table[
        (
            sku_table[
                "실제검수수량"
            ]
            >= min_inspection_qty
        )
        &
        (
            sku_table[
                "기초재고"
            ]
            +
            sku_table[
                "입고수량"
            ]
            >= min_available_qty
        )
    ].copy()

    sku_table = sku_table.sort_values(
        [
            "불량률",
            "불량수량",
            "실제검수수량",
        ],
        ascending=[
            False,
            False,
            False,
        ],
    ).head(
        int(top_n)
    ).reset_index(
        drop=True
    )

    if sku_table.empty:

        st.info(
            "현재 조건에서 표시할 SKU가 없습니다."
        )

    else:

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
            ]
        ].copy()

        show_table(
            sku_display,
            height=600,
        )

        st.caption(
            "불량률 = 불량수량 ÷ 실제 검수수량 / "
            "실제 검수수량 = 출고수량 + 불량수량"
        )


# ============================================================
# 끝
# ============================================================

st.divider()


st.caption(
    "※ 입고/출고/불량 수량은 원본 거래 수량을 직접 집계하며, "
    "상품마스터 연결은 수량 집계 이후의 기준정보 매핑에만 사용합니다."
)
