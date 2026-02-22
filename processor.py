import pandas as pd

def split_id_nickname(text):
    text = str(text)
    if "(" in text and ")" in text:
        id_part, nick_part = text.split("(", 1)
        nick_part = nick_part.rstrip(")")
    else:
        id_part = text
        nick_part = ""
    return id_part.strip(), nick_part.strip()


def classify_heart(id_val):
    id_val = str(id_val)
    if "@ka" in id_val:
        return "일반"
    if "@" in id_val:
        return "제휴"
    return "일반"


def process_dataframe(df):

    # 필요한 컬럼만 사용
    df = df.copy()
    df = df[["후원 아이디(닉네임)", "후원하트", "참여BJ"]]

    # 아이디 / 닉네임 분리
    df[["후원아이디", "닉네임"]] = df["후원 아이디(닉네임)"].apply(
        lambda x: pd.Series(split_id_nickname(x))
    )

    # 하트 숫자 정리
    df["후원하트"] = pd.to_numeric(df["후원하트"], errors="coerce").fillna(0)
    df.loc[df["후원하트"] < 0, "후원하트"] = 0

    # 하트 타입 구분
    df["하트구분"] = df["후원아이디"].apply(classify_heart)

    result = {}

    for bj, bj_df in df.groupby("참여BJ"):

        # ==============================
        # 1️⃣ 아이디 기준 합산
        # ==============================

        # 아이디 + 닉네임별 합계
        nick_sum = (
            bj_df.groupby(["후원아이디", "닉네임"])["후원하트"]
            .sum()
            .reset_index()
        )

        # 각 아이디에서 가장 하트 많이 받은 닉네임 선택
        idx = nick_sum.groupby("후원아이디")["후원하트"].idxmax()
        representative = nick_sum.loc[idx]

        # 아이디 기준 총합
        total_sum = (
            bj_df.groupby("후원아이디")["후원하트"]
            .sum()
            .reset_index()
        )

        merged = pd.merge(total_sum, representative[["후원아이디", "닉네임"]],
                          on="후원아이디", how="left")

        # 하트 타입 다시 붙이기
        merged["하트구분"] = merged["후원아이디"].apply(classify_heart)

        # ==============================
        # 2️⃣ 정산용 정렬 (일반 위 / 제휴 아래)
        # ==============================

        normal = merged[merged["하트구분"] == "일반"].sort_values("후원하트", ascending=False)
        partner = merged[merged["하트구분"] == "제휴"].sort_values("후원하트", ascending=False)

        settlement_view = pd.concat([normal, partner])

        # ==============================
        # 3️⃣ BJ용 정렬 (전체 통합 내림차순)
        # ==============================

        bj_view = merged.sort_values("후원하트", ascending=False)

        result[bj] = {
            "정산용": settlement_view.reset_index(drop=True),
            "BJ용": bj_view.reset_index(drop=True)
        }

    return result

