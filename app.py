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
# 📅 날짜 prefix (파일명 우선 → 없으면 데이터 최솟값)
# ==================================================
def extract_prefix_from_filename(files):
    """
    업로드 파일명 앞에 'MM.DD' 형태가 있으면 그 값을 prefix로 사용
    예) '01.01 하꿍_내역.csv' -> '01.01'
    """
    for f in files:
        stem = Path(f.name).stem
        m = re.match(r"^(\d{2}\.\d{2})", stem)
        if m:
            return m.group(1)
    return None


def extract_earliest_date_prefix(df):
    """
    데이터 내 '후원시간' 계열 컬럼에서 가장 이른 날짜를 찾아 'MM.DD' 반환
    """
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

# 업로드 파일이 1개일 때만 날짜 사용
if len(uploaded_files) == 1:
    prefix = extract_prefix_from_filename(uploaded_files)
    if not prefix:
        prefix = extract_earliest_date_prefix(merged)
else:
    prefix = None


# ==================================================
# 📊 웹 1차 요약표 (참여BJ별 일반/제휴/총합)
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

    # 너 원본 로직과 호환되게 "후원+아이디"만 찾도록 완화 (닉네임 포함 조건 제거)
    col_idnick = next((c for c in tmp.columns if "후원" in c and "아이디" in c), None)
    col_heart = next((c for c in tmp.columns if "후원" in c and "하트" in c), None)
    col_bj = next((c for c in tmp.columns if "참여" in c and "BJ" in c), None)

    if not (col_idnick and col_heart and col_bj):
        st.warning("필수 컬럼(후원아이디/후원하트/참여BJ)을 찾지 못했습니다.")
    else:
        tmp[col_heart] = pd.to_numeric(tmp[col_heart], errors="coerce").fillna(0)
        tmp.loc[tmp[col_heart] < 0, col_heart] = 0

        # '(닉네임)' 같이 붙은 포맷이면 괄호 제거
        tmp["후원아이디"] = (
            tmp[col_idnick]
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

        # 화면 표시용 천단위 콤마 (데이터 자체는 문자열로 바뀜: 표시만 목적)
        for c in ["일반", "제휴", "총합"]:
            pivot[c] = pivot[c].apply(lambda x: f"{int(x):,}")

        st.subheader("요약_참여BJ_총계")
        st.dataframe(pivot.reset_index(drop=True), hide_index=True, use_container_width=True)

except Exception as e:
    st.warning(f"요약표 생성 중 오류: {e}")


# ==================================================
# 📁 BJ별 엑셀 다운로드
# ==================================================
result = process_dataframe(merged)

if not result:
    st.error("집계 결과가 없습니다.")
    st.stop()


def make_excel(df: pd.DataFrame, bj_name: str) -> BytesIO:
    """
    processor가 반환한 df(정산용/BJ용)를 받아 엑셀(BytesIO)로 만들어 반환
    df에는 최소 '후원아이디', '닉네임', '후원하트' 컬럼이 있어야 한다.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "정산표"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    total = int(pd.to_numeric(df["후원하트"], errors="coerce").fillna(0).sum())

    ws["A1"] = ""
    ws["B1"] = bj_name
    ws["C1"] = total

    ws["A2"] = "후원아이디"
    ws["B2"] = "닉네임"
    ws["C2"] = "후원하트"

    for col in ["A", "B", "C"]:
        ws[f"{col}2"].border = border
        ws[f"{col}2"].alignment = Alignment(horizontal="center")

    row = 3
    for _, r in df.iterrows():
        ws.cell(row=row, column=1, value=str(r["후원아이디"]))
        ws.cell(row=row, column=2, value=str(r["닉네임"]))
        heart = pd.to_numeric(r["후원하트"], errors="coerce")
        heart = 0 if pd.isna(heart) else int(heart)
        heart = max(heart, 0)
        ws.cell(row=row, column=3, value=heart).number_format = "#,##0"
        row += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 12

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


st.success("집계 완료")

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
