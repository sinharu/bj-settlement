import pandas as pd

def _read_csv_any(path: str) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return pd.read_csv(path, encoding=enc)
        except Exception:
            pass
    return pd.read_csv(path, encoding="utf-8", errors="replace")

def load_file(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        return _read_csv_any(path)
    return pd.read_excel(path)

def split_id_nickname(text):
    """
    'id(닉네임)' -> ('id', '닉네임')
    괄호가 없으면 ('원문', '')
    """
    try:
        s = "" if pd.isna(text) else str(text)
        if "(" in s and ")" in s:
            _id, _nick = s.split("(", 1)
            return _id.strip(), _nick.rstrip(")").strip()
        return s.strip(), ""
    except Exception:
        return str(text), ""

def heart_type(uid: str) -> int:
    """
    0 = 일반하트
    1 = 제휴하트
    규칙:
    - '@' 있으면 제휴
    - 단 '@ka' 포함 시 예외로 일반
    """
    uid = str(uid)
    if "@ka" in uid:
        return 0
    if "@" in uid:
        return 1
    return 0

def process_dataframe(df: pd.DataFrame) -> dict:
    """
    원본 구조 (고정)
    A열: 일자/시간 (무시)
    B열: 후원아이디(닉네임)
    C열: 후원하트
    D열: 참여BJ (이걸로 파일 분리)

    출력 구조
    A열: 후원아이디
    B열: 닉네임
    C열: 후원하트 (닉네임 기준 합산)

    정렬 규칙
    1. 일반하트 먼저
    2. 제휴하트 아래
    3. 각 그룹 내에서 후원하트 내림차순
    """

    if df is None or df.empty:
        return {}

    if df.shape[1] < 4:
        raise ValueError(
            "컬럼 부족: A=시간(무시), B=후원아이디(닉), C=후원하트, D=참여BJ 필요"
        )

    df = df.copy()

    # 컬럼 위치 기준
    idnick_col = df.iloc[:, 1]   # B열
    heart_col  = df.iloc[:, 2]   # C열
    bj_col     = df.iloc[:, 3]   # D열

    df["_BJ"] = bj_col.astype(str).str.strip()

    # ID / 닉네임 분리
    df[["_ID", "_NICK"]] = idnick_col.apply(
        lambda x: pd.Series(split_id_nickname(x))
    )

    # 하트 숫자화 + 음수 제거
    df["_HEART"] = pd.to_numeric(heart_col, errors="coerce").fillna(0)
    df.loc[df["_HEART"] < 0, "_HEART"] = 0

    result = {}

    for bj in df["_BJ"].dropna().unique():
        bj_df = df[df["_BJ"] == bj]
        if bj_df.empty:
            continue

        grouped = (
            bj_df
            .groupby(["_NICK"], as_index=False)
            .agg(
                후원하트=("_HEART", "sum"),
                후원아이디=("_ID", "first")
            )
        )

        # 일반 / 제휴 구분
        grouped["_TYPE"] = grouped["후원아이디"].apply(heart_type)
        # 0 = 일반, 1 = 제휴

        # 정렬: 일반 → 제휴, 각 그룹 내 하트 내림차순
        grouped = grouped.sort_values(
            by=["_TYPE", "후원하트"],
            ascending=[True, False]
        )

        out = (
            grouped
            .rename(columns={"_NICK": "닉네임"})
            [["후원아이디", "닉네임", "후원하트"]]
            .reset_index(drop=True)
        )

        result[bj] = out

    return result

