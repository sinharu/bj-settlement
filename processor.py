import unicodedata
from decimal import Decimal, InvalidOperation

import pandas as pd


FORMAT_HEART = "하트"
FORMAT_KRW = "원화"

STANDARD_COLUMNS = [
    "형식", "참여BJ", "회차", "후원시간", "날짜", "시간",
    "아이디", "닉네임", "후원하트", "구분", "후원종류",
    "방송일", "방송회차",
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
def _find_ledger_header(df: pd.DataFrame) -> tuple[str, int] | None:
    if df.columns.empty or not isinstance(df.columns[0], str):
        return None
    title = df.columns[0].strip()
    suffix = " 총 후원 내역"
    if not title.endswith(suffix):
        return None
    bj_name = normalize_bj_name(title[:-len(suffix)])
    if not bj_name:
        return None

    required = {"후원 일시", "플랫폼", "후원자 닉네임", "총 후원 금액"}
    for position, row in enumerate(df.itertuples(index=False, name=None)):
        if required.issubset({value.strip() for value in row if isinstance(value, str)}):
            return bj_name, position
    return None


def _ledger_amount(value, label: str) -> int:
    """장부 금액은 누락/잘못된 값을 0으로 바꾸거나 소수점을 버리지 않는다."""
    try:
        amount = Decimal(str(value).strip())
    except InvalidOperation:
        amount = Decimal("NaN")
    if (
        not amount.is_finite()
        or amount < 0
        or amount != amount.to_integral_value()
        or amount > 2**63 - 1
    ):
        raise ValueError(f"후원 장부 {label}: 유효한 0 이상의 원화 정수가 아닙니다 (값: {value}).")
    return int(amount)


def _ledger_to_krw(df: pd.DataFrame, bj_name: str, header_position: int):
    columns = {
        value.strip(): position
        for position, value in enumerate(df.iloc[header_position])
        if isinstance(value, str)
    }
    required = {
        "번호", "후원 일시", "플랫폼", "후원자 닉네임",
        "웹후원 금액", "계좌후원 금액", "총 후원 금액",
    }
    missing = required.difference(columns)
    if missing:
        raise ValueError(f"후원 장부 필수 열이 없습니다: {', '.join(sorted(missing))}")

    metadata = {
        row[0].strip(): row[1]
        for row in df.iloc[:header_position].itertuples(index=False, name=None)
        if isinstance(row[0], str)
    }
    broadcast_date = pd.to_datetime(metadata.get("방송일"), errors="coerce")
    broadcast_date = None if pd.isna(broadcast_date) else broadcast_date.date()
    broadcast_round = normalize_bj_name(metadata.get("방송 회차명")) or None
    round_column = df.columns.get_loc("업로드회차") if "업로드회차" in df.columns else None

    totals = {"웹후원 합계": 0, "계좌후원 합계": 0, "총 후원 합계": 0}
    footer_totals = []
    records = []
    for excel_row, row in enumerate(
        df.iloc[header_position + 1:].itertuples(index=False, name=None),
        start=header_position + 3,
    ):
        label = row[0].strip() if isinstance(row[0], str) else ""
        if label in totals:
            footer_totals.append((label, _ledger_amount(row[1], f"엑셀 {excel_row}행 {label}")))
            continue
        if pd.isna(pd.to_numeric(row[columns["번호"]], errors="coerce")):
            continue

        web = _ledger_amount(row[columns["웹후원 금액"]], f"엑셀 {excel_row}행 웹후원 금액")
        bank = _ledger_amount(row[columns["계좌후원 금액"]], f"엑셀 {excel_row}행 계좌후원 금액")
        total = _ledger_amount(row[columns["총 후원 금액"]], f"엑셀 {excel_row}행 총 후원 금액")
        if web + bank != total:
            raise ValueError(
                f"후원 장부 엑셀 {excel_row}행 금액 불일치: 웹후원 금액 {web:,}원 + "
                f"계좌후원 금액 {bank:,}원 = {web + bank:,}원, "
                f"총 후원 금액 {total:,}원 (차이 {abs(web + bank - total):,}원)."
            )
        totals["웹후원 합계"] += web
        totals["계좌후원 합계"] += bank
        totals["총 후원 합계"] += total

        platform = normalize_bj_name(row[columns["플랫폼"]])
        web_type = platform if platform and platform != "계좌후원" else "웹후원"
        for amount, donation_type in [(web, web_type), (bank, "계좌후원")]:
            if amount > 0:
                records.append({
                    "스트리머": bj_name,
                    "닉네임": row[columns["후원자 닉네임"]],
                    "종류": donation_type,
                    "개수": amount,
                    "시간": row[columns["후원 일시"]],
                    "업로드회차": row[round_column] if round_column is not None else None,
                })

    for label, reported in footer_totals:
        actual = totals[label]
        if reported != actual:
            raise ValueError(
                f"후원 장부 {label} 불일치: 하단 합계 {reported:,}원, "
                f"데이터 합계 {actual:,}원 (차이 {abs(reported - actual):,}원)."
            )

    return (
        pd.DataFrame(records, columns=["스트리머", "닉네임", "종류", "개수", "시간", "업로드회차"]),
        broadcast_date,
        broadcast_round,
    )


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

    if _find_ledger_header(df) is not None:
        return FORMAT_KRW

    return None


def to_standard(df: pd.DataFrame) -> pd.DataFrame | None:
    """후원 수량의 단위를 유지하면서 지원하는 입력 형식을 공통 열로 변환한다."""
    donation_format = detect_format(df)
    if donation_format is None:
        return None

    df = df.reset_index(drop=True).copy()
    broadcast_date = None
    broadcast_round = None
    if donation_format == FORMAT_KRW:
        ledger_header = _find_ledger_header(df)
        if ledger_header is not None:
            df, broadcast_date, broadcast_round = _ledger_to_krw(df, *ledger_header)
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
    standard["방송일"] = broadcast_date
    standard["방송회차"] = broadcast_round

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
        standard = df.copy()
        for column in ["방송일", "방송회차"]:
            if column not in standard.columns:
                standard[column] = None
        return standard
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
