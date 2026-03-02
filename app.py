import re
from pathlib import Path
from io import BytesIO

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side

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
# 📦 엑셀 공통 유틸
# ==================================================
thin = Side(style="thin")
all_border = Border(left=thin, right=thin, top=thin, bottom=thin)

def apply_border(ws):
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value not in (None, ""):
                cell.border = all_border

def auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 18), 45)

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
# 📥 파일 읽기
# ==================================================
dfs = []
for f in uploaded_files:
    if f.name.lower().endswith(".csv"):
        dfs.append(pd.read_csv(f))
    else:
        dfs.append(pd.read_excel(f))

merged = pd.concat(dfs, ignore_index=True)

# ==================================================
# 📅 파일 1개 업로드 시 날짜 prefix
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
    if "@ka" in str(user_id):
        return "일반"
    if "@" in str(user_id):
        return "제휴"
    return "일반"

tmp = merged.copy()
col_id = next((c for c in tmp.columns if "후원" in c and "아이디" in c), None)
col_heart = next((c for c in tmp.columns if "후원" in c and "하트" in c), None)
col_bj = next((c for c in tmp.columns if "참여" in c and "BJ" in c), None)

if col_id and col_heart and col_bj:
    tmp[col_heart] = pd.to_numeric(tmp[col_heart], errors="coerce").fillna(0)
    tmp.loc[tmp[col_heart] < 0, col_heart] = 0

    tmp["후원아이디"] = tmp[col_id].astype(str).str.replace(r"\(.*\)", "", regex=True).str.strip()
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

# ==================================================
# 📁 BJ별 엑셀
# ==================================================
result = process_dataframe(merged)

def make_excel(df: pd.DataFrame, bj_name: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "정산표"

    total = int(pd.to_numeric(df["후원하트"], errors="coerce").fillna(0).sum())

    ws.cell(row=1, column=2, value=bj_name)
    total_cell = ws.cell(row=1, column=3, value=total)
    total_cell.number_format = "#,##0"

    ws.append(["후원아이디", "닉네임", "후원하트"])

    for _, r in df.iterrows():
        row = ws.max_row + 1
        ws.cell(row=row, column=1, value=r["후원아이디"])
        ws.cell(row=row, column=2, value=r["닉네임"])
        heart = int(pd.to_numeric(r["후원하트"], errors="coerce") or 0)
        heart_cell = ws.cell(row=row, column=3, value=heart)
        heart_cell.number_format = "#,##0"

    auto_width(ws)
    apply_border(ws)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ==================================================
# 📥 다운로드
# ==================================================
st.success("집계 완료")

if len(uploaded_files) > 1:
    total_file = make_total_excel(merged)
    total_filename = f"{prefix}_총합산.xlsx" if prefix else "총합산.xlsx"

    st.download_button(
        label=f"{total_filename} 다운로드",
        data=total_file,
        file_name=total_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

for bj, views in result.items():
    filename1 = f"{prefix}_{bj}_정산용.xlsx" if prefix else f"{bj}_정산용.xlsx"
    filename2 = f"{prefix}_{bj}_BJ용.xlsx" if prefix else f"{bj}_BJ용.xlsx"

    st.subheader(bj)

    st.download_button(
        label=f"{filename1} 다운로드",
        data=make_excel(views["정산용"], bj),
        file_name=filename1,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label=f"{filename2} 다운로드",
        data=make_excel(views["BJ용"], bj),
        file_name=filename2,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
