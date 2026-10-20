import unicodedata

import pandas as pd


FORMAT_HEART = "하트"
FORMAT_KRW = "원화"

STANDARD_COLUMNS = [
    "형식", "참여BJ", "회차", "후원시간", "날짜", "시간",
    "아이디", "닉네임", "후원하트", "구분", "후원종류",
]


def normalize_bj_name(value) -> str:
    """같은 BJ가 앞뒤 공백 차이로 분리되지 않도록 이름을 표준화한다."""
    if pd.isna(value):
        return ""
    return unicodedata.normalize("NFC", str(value)).strip()


def parse_donation_times(series):
    text = series.astype(str).str.strip()
    text = text.str.replace("오전", "AM", regex=False).str.replace("오후", "PM", regex=False)
    parsed = pd.to_datetime(text, errors="coerce", format="mixed")

    missing = parsed.isna()
    if missing.any():
        numeric = pd.to_numeric(series[missing], errors="coerce")
        excel_dates = pd.to_datetime(numeric, errors="coerce", unit="D", origin="1899-12-30")
        parsed.loc[missing] = excel_dates

    return parsed


# ==========================================
# 🔹 ID / 닉네임 분리
# ==========================================
def split_id_nickname(text):
    text = str(text)

    if "(" in text and ")" in text:
        id_part, nick_part = text.split("(", 1)
        nick_part = nick_part.rstrip(")")
    else:
        id_part = text
        nick_part = ""

    return id_part.strip(), nick_part.strip()


# ==========================================
# 🔹 하트 구분
# ==========================================
def classify_heart(user_id: str) -> str:
    s = str(user_id)

    if "@ka" in s:
        return "일반"
    if "@" in s:
        return "제휴"
    return "일반"


# ==========================================
# 🔹 전처리 + 표준화
# ==========================================
def detect_format(df: pd.DataFrame) -> str | None:
    """필수 열 이름으로 하트/원화 후원 형식을 구분한다."""
    columns = [c for c in df.columns if isinstance(c, str)]
    if {"닉네임", "스트리머", "종류", "개수", "시간"}.issubset(
        {c.strip() for c in columns}
    ):
        return FORMAT_KRW

    if all(
        any(first in c and second in c for c in columns)
        for first, second in [("후원", "아이디"), ("후원", "하트"), ("참여", "BJ")]
    ):
        return FORMAT_HEART

    return None


def to_standard(df: pd.DataFrame) -> pd.DataFrame | None:
    """후원 수량의 단위를 유지하면서 두 입력 형식을 공통 열로 변환한다."""
    donation_format = detect_format(df)
    if donation_format is None:
        return None

    df = df.reset_index(drop=True).copy()
    if donation_format == FORMAT_KRW:
        columns = {c.strip(): c for c in df.columns if isinstance(c, str)}
        if "등록" in columns:
            registered = df[columns["등록"]].fillna("").astype(str).str.strip()
            df = df.loc[registered == "등록"].copy()

        standard = pd.DataFrame(index=df.index)
        nickname = df[columns["닉네임"]]
        if "아이디" in columns:
            nickname = nickname.fillna(df[columns["아이디"]])
        standard["닉네임"] = nickname.map(normalize_bj_name)
        standard["아이디"] = standard["닉네임"]
        standard["구분"] = "일반"
        standard["후원종류"] = df[columns["종류"]].fillna("").astype(str).str.strip()
        col_bj = columns["스트리머"]
        col_amount = columns["개수"]
        col_time = columns["시간"]
    else:
        columns = [c for c in df.columns if isinstance(c, str)]
        col_idnick = next(c for c in columns if "후원" in c and "아이디" in c)
        col_amount = next(c for c in columns if "후원" in c and "하트" in c)
        col_bj = next(c for c in columns if "참여" in c and "BJ" in c)
        col_time = next((c for c in columns if "후원" in c and "시간" in c), None)
        standard = pd.DataFrame(
            df[col_idnick].map(split_id_nickname).tolist(),
            index=df.index,
            columns=["아이디", "닉네임"],
        )
        standard["구분"] = standard["아이디"].map(classify_heart)
        standard["후원종류"] = ""

    # 호환 열 이름을 유지하며 하트 수 또는 원화 금액을 정수로 보관한다.
    standard["후원하트"] = (
        pd.to_numeric(df[col_amount], errors="coerce").fillna(0).clip(lower=0).astype("int64")
    )
    standard["형식"] = donation_format
    standard["참여BJ"] = df[col_bj].map(normalize_bj_name)
    standard["회차"] = df["업로드회차"] if "업로드회차" in df.columns else None

    if col_time is not None:
        standard["후원시간"] = parse_donation_times(df[col_time])
        standard["날짜"] = standard["후원시간"].dt.date
        standard["시간"] = standard["후원시간"].dt.time
    else:
        standard["후원시간"] = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
        standard["날짜"] = None
        standard["시간"] = None

    return standard.loc[standard["참여BJ"] != "", STANDARD_COLUMNS].reset_index(drop=True)


def clean_and_prepare(df: pd.DataFrame):
    if "형식" in df.columns:
        return df.copy()
    return to_standard(df)


# ==========================================
# 🔹 메인 집계
# ==========================================
def process_dataframe(df: pd.DataFrame):

    df = clean_and_prepare(df)
    if df is None or df.empty:
        return None

    result = {}

    for bj, bj_df in df.groupby("참여BJ"):

        # =====================================
        # 1️⃣ 아이디 + 닉네임별 합산
        # =====================================
        nick_sum = (
            bj_df.groupby(["아이디", "닉네임"])["후원하트"]
            .sum()
            .reset_index()
        )

        # 각 아이디에서 가장 하트 많이 받은 닉네임 선택
        idx = nick_sum.groupby("아이디")["후원하트"].idxmax()
        representative = nick_sum.loc[idx]

        # 아이디 기준 총합
        total_sum = (
            bj_df.groupby("아이디")["후원하트"]
            .sum()
            .reset_index()
        )

        merged = pd.merge(
            total_sum,
            representative[["아이디", "닉네임"]],
            on="아이디",
            how="left"
        )

        # 원화 닉네임에 @가 있어도 표준화 단계에서 정한 일반 구분을 유지한다.
        merged["구분"] = merged["아이디"].map(bj_df.groupby("아이디")["구분"].first())

        # =====================================
        # 2️⃣ 정산용 정렬
        # 일반 위 / 제휴 아래 / 각 그룹 내 내림차순
        # =====================================
        normal = merged[merged["구분"] == "일반"].sort_values("후원하트", ascending=False)
        partner = merged[merged["구분"] == "제휴"].sort_values("후원하트", ascending=False)

        settlement_view = pd.concat([normal, partner]).reset_index(drop=True)

        # =====================================
        # 3️⃣ BJ용 정렬 (전체 통합 내림차순)
        # =====================================
        bj_view = merged.sort_values("후원하트", ascending=False).reset_index(drop=True)

        result[bj] = {
            "정산용": settlement_view,
            "BJ용": bj_view,
            "전체로그": bj_df[STANDARD_COLUMNS].reset_index(drop=True)
        }

    return result
