import pandas as pd


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
def clean_and_prepare(df: pd.DataFrame):

    df = df.copy()

    # 컬럼 자동 탐색
    col_idnick = next((c for c in df.columns if "후원" in c and "아이디" in c), None)
    col_heart = next((c for c in df.columns if "후원" in c and "하트" in c), None)
    col_bj = next((c for c in df.columns if "참여" in c and "BJ" in c), None)
    col_time = next((c for c in df.columns if "후원" in c and "시간" in c), None)

    if not all([col_idnick, col_heart, col_bj]):
        return None

    # 아이디 / 닉네임 분리
    df[["아이디", "닉네임"]] = df[col_idnick].apply(
        lambda x: pd.Series(split_id_nickname(x))
    )

    # 하트 숫자 정리
    df["후원하트"] = pd.to_numeric(df[col_heart], errors="coerce").fillna(0)
    df.loc[df["후원하트"] < 0, "후원하트"] = 0

    # 하트 타입
    df["구분"] = df["아이디"].apply(classify_heart)

    # 날짜/시간 처리
    if col_time:
        df["후원시간"] = parse_donation_times(df[col_time])
        df["날짜"] = df["후원시간"].dt.date
        df["시간"] = df["후원시간"].dt.time
    else:
        df["날짜"] = None
        df["시간"] = None

    if "업로드회차" in df.columns:
        df["회차"] = df["업로드회차"]
    else:
        df["회차"] = None

    df["참여BJ"] = df[col_bj]

    return df[["참여BJ", "회차", "날짜", "시간", "아이디", "닉네임", "후원하트", "구분"]]


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

        merged["구분"] = merged["아이디"].apply(classify_heart)

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
            "전체로그": bj_df.reset_index(drop=True)
        }

    return result
