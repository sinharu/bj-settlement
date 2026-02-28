import re
from pathlib import Path
from io import BytesIO

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment

from processor import process_dataframe


st.set_page_config(page_title="BJ 하트 집계", layout="centered")

# ==================================================
# 🔐 비밀번호 게이트
# ==================================================
def check_password():
    def password_entered():
        if st.session_state.get("password", "") == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            st.session_state.pop("password", None)
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("비밀번호를 입력하세요", type="password", key="password", on_change=password_entered)
        return False

    if not st.session_state["password_correct"]:
        st.text_input("비밀번호를 입력하세요", type="password", key="password", on_change=password_entered)
        st.error("비밀번호가 틀렸습니다.")
        return False

    return True


if not check_password():
    st.stop()


# ==================================================
# 📌 화면 시작
# ==================================================
st.title("BJ 하트 집계 (BJ 전달용)")
st.caption("CSV / XLSX 업로드 → 웹 요약표 확인 → BJ별 엑셀 다운로드")

uploaded_files = st.file_uploader(
    "CSV 또는 XLSX 파일을 업로드하세요",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("파일을 업로드하면 집계 결과가 표시됩니다.")
    st.stop()


# ==================================================
# 📅 날짜 prefix (파일 1개일 때만 적용)
# ==================================================
def extract_prefix_from_filename(files):
    for f in files:
        stem = Path(f.name).stem
        m = re.match(r"^(\d{2}\.\d{2})", stem)
        if m:
            return m.group(1)
    return None


def extract_earliest_date_prefix(df):
    col_time = next((c for c in df.columns if "후원" in c and "시간" in c), None)
    if not col_time:
        return None
    tmp = df[[col_time]].copy()
    tmp[col_time] = pd.to_datetime(tmp[col_time], errors="coerce")
    min_dt = tmp[col_time].min()
    if pd.isna(min_dt):
        return None
    return min_dt.strftime("%m.%d")


# ==================================================
# 📥 파일 읽기
# ==================================================
dfs = []

for f in uploaded_files:
    try:
        if f.name.lower().endswith(".csv"):
            dfs.append(pd.read_csv(f))
        else:
            dfs.append(pd.read_excel(f))
    except Exception as e:
        st.error(f"{f.name} 읽기 실패: {e}")

if not dfs:
    st.error("읽을 수 있는 파일이 없습니다.")
    st.stop()

merged = pd.concat(dfs, ignore_index=True)

if len(uploaded_files) == 1:
    prefix = extract_prefix_from_filename(uploaded_files)
    if not prefix:
        prefix = extract_earliest_date_prefix(merged)
else:
    prefix = None


# ==================================================
# 📊 웹 요약표
# ==================================================
def classify_heart_type(user_id: str) -> str:
    s = str(user_id)
    if "@ka" in s:
        return "일반"
    if "@" in s:
        return "제휴"
    return "일반"


try:
    tmp = merged.copy()

    col_id = next((c for c in tmp.columns if "후원" in c and "아이디" in c), None)
    col_heart = next((c for c in tmp.columns if "후원" in c and "하트" in c), None)
    col_bj = next((c for c in tmp.columns if "참여" in c and "BJ" in c), None)

    if not (col_id and col_heart and col_bj):
        st.warning("필수 컬럼을 찾지 못했습니다.")
    else:
        tmp[col_heart] = pd.to_numeric(tmp[col_heart], errors="coerce").fillna(0)
        tmp.loc[tmp[col_heart] < 0, col_heart] = 0

        tmp["후원아이디"] = (
            tmp[col_id]
            .astype(str)
            .str.replace(r"\(.*\)", "", regex=True)
            .str.strip()
        )

        tmp["구분"] = tmp["후원아이디"].apply(classify_heart_type)

        pivot = (
            tmp.groupby([col_bj, "구분"])[col_heart]
            .sum()
            .unstack(fill_value=0)
            .reset_index()
        )

        if "일반" not in pivot.columns:
            pivot["일반"] = 0
        if "제휴" not in pivot.columns:
            pivot["제휴"] = 0

        pivot["총합"] = pivot["일반"] + pivot["제휴"]
        pivot = pivot.rename(columns={col_bj: "참여BJ"})
        pivot = pivot[["참여BJ", "일반", "제휴", "총합"]].sort_values("총합", ascending=False)

        for c in ["일반", "제휴", "총합"]:
            pivot[c] = pivot[c].apply(lambda x: f"{int(x):,}")

        st.subheader("요약_참여BJ_총계")
        st.dataframe(pivot.reset_index(drop=True), hide_index=True, use_container_width=True)

except Exception as e:
    st.warning(f"요약표 생성 중 오류: {e}")


# ==================================================
# 📁 BJ별 다운로드
# ==================================================
result = process_dataframe(merged)

if not result:
    st.error("집계 결과가 없습니다.")
    st.stop()


def make_excel(df: pd.DataFrame, bj_name: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "정산표"

    total = int(pd.to_numeric(df["후원하트"], errors="coerce").fillna(0).sum())

    ws.append(["", bj_name, total])
    ws.append(["후원아이디", "닉네임", "후원하트"])

    for _, r in df.iterrows():
        heart = pd.to_numeric(r["후원하트"], errors="coerce")
        heart = 0 if pd.isna(heart) else int(heart)
        ws.append([r["후원아이디"], r["닉네임"], heart])

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ==================================================
# 📦 총합산 파일 (여러 파일 업로드 시)
# ==================================================
def make_total_excel(df: pd.DataFrame) -> BytesIO:

    wb = Workbook()
    wb.remove(wb.active)

    tmp = df.copy()

    col_time = next((c for c in tmp.columns if "후원" in c and "시간" in c), None)
    col_idnick = next((c for c in tmp.columns if "후원" in c and "아이디" in c), None)
    col_heart = next((c for c in tmp.columns if "후원" in c and "하트" in c), None)
    col_bj = next((c for c in tmp.columns if "참여" in c and "BJ" in c), None)

    if not all([col_time, col_idnick, col_heart, col_bj]):
        return None

    tmp[col_time] = pd.to_datetime(tmp[col_time], errors="coerce")
    tmp["날짜"] = tmp[col_time].dt.date
    tmp["시간"] = tmp[col_time].dt.time
    tmp[col_heart] = pd.to_numeric(tmp[col_heart], errors="coerce").fillna(0)

    def split_id_nickname(text):
        text = str(text)
        if "(" in text and ")" in text:
            id_part, nick_part = text.split("(", 1)
            nick_part = nick_part.rstrip(")")
        else:
            id_part = text
            nick_part = ""
        return id_part.strip(), nick_part.strip()

    tmp[["아이디", "닉네임"]] = tmp[col_idnick].apply(
        lambda x: pd.Series(split_id_nickname(x))
    )

    def classify(x):
        if "@ka" in x:
            return "일반"
        if "@" in x:
            return "제휴"
        return "일반"

    tmp["구분"] = tmp["아이디"].apply(classify)

    # ==========================
    # 1️⃣ 일자별 집계
    # ==========================
    ws1 = wb.create_sheet("일자별집계")
    ws1.append(["날짜", "BJ", "일반", "제휴", "총합"])

    s1 = (
        tmp.groupby(["날짜", col_bj, "구분"])[col_heart]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    if "일반" not in s1.columns: s1["일반"] = 0
    if "제휴" not in s1.columns: s1["제휴"] = 0

    s1["총합"] = s1["일반"] + s1["제휴"]

    for _, r in s1.iterrows():
        ws1.append([
            r["날짜"],
            r[col_bj],
            int(r["일반"]),
            int(r["제휴"]),
            int(r["총합"])
        ])

    # 넓게 설정
    ws1.column_dimensions["A"].width = 14
    ws1.column_dimensions["B"].width = 10
    ws1.column_dimensions["C"].width = 12
    ws1.column_dimensions["D"].width = 12
    ws1.column_dimensions["E"].width = 12

    ws1.sheet_view.zoomScale = 110


    # ==========================
    # 2️⃣ 전체 총합
    # ==========================
    ws2 = wb.create_sheet("총합")
    ws2.append(["BJ", "일반", "제휴", "총합"])

    s2 = (
        tmp.groupby([col_bj, "구분"])[col_heart]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    if "일반" not in s2.columns: s2["일반"] = 0
    if "제휴" not in s2.columns: s2["제휴"] = 0

    s2["총합"] = s2["일반"] + s2["제휴"]

    for _, r in s2.iterrows():
        ws2.append([
            r[col_bj],
            int(r["일반"]),
            int(r["제휴"]),
            int(r["총합"])
        ])

    ws2.column_dimensions["A"].width = 10
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 12
    ws2.column_dimensions["D"].width = 12

    ws2.sheet_view.zoomScale = 110


    # ==========================
    # 3️⃣ BJ별 상세
    # ==========================
    for bj in tmp[col_bj].unique():

        ws = wb.create_sheet(str(bj))
        sub = tmp[tmp[col_bj] == bj]

        일반합 = sub[sub["구분"] == "일반"][col_heart].sum()
        제휴합 = sub[sub["구분"] == "제휴"][col_heart].sum()
        총합 = 일반합 + 제휴합

        ws["A1"] = "총하트"
        ws["B1"] = int(총합)

        ws["D1"] = "일반하트"
        ws["E1"] = int(일반합)

        ws["G1"] = "제휴하트"
        ws["H1"] = int(제휴합)

        ws["B1"].number_format = "#,##0"
        ws["E1"].number_format = "#,##0"
        ws["H1"].number_format = "#,##0"

        ws.append([])
        ws.append(["날짜", "시간", "아이디", "닉네임", "하트", "구분"])

        for _, r in sub.iterrows():
            ws.append([
                r["날짜"],
                r["시간"],
                r["아이디"],
                r["닉네임"],
                int(r[col_heart]),
                r["구분"]
            ])

        # 매우 넓게
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 20
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 12
        ws.column_dimensions["F"].width = 12

        ws.sheet_view.zoomScale = 110

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    return bio

st.success("집계 완료")

if len(uploaded_files) > 1:
    total_file = make_total_excel(merged)
    if total_file:
        st.download_button(
            label="총합산.xlsx 다운로드",
            data=total_file,
            file_name="총합산.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

for bj, views in result.items():
    settlement_df = views["정산용"]
    bj_df = views["BJ용"]

    st.subheader(bj)

    filename1 = f"{prefix}_{bj}_정산용.xlsx" if prefix else f"{bj}_정산용.xlsx"
    filename2 = f"{prefix}_{bj}_BJ용.xlsx" if prefix else f"{bj}_BJ용.xlsx"

    st.download_button(
        label=f"{filename1} 다운로드",
        data=make_excel(settlement_df, bj),
        file_name=filename1,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label=f"{filename2} 다운로드",
        data=make_excel(bj_df, bj),
        file_name=filename2,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
