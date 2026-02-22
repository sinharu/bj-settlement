import streamlit as st
import pandas as pd
from processor import process_dataframe
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="BJ 하트 집계", layout="centered")

# =========================
# 🔐 비밀번호 게이트
# =========================
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

# =========================
# 기존 화면
# =========================
st.title("BJ 하트 집계 (BJ 전달용)")
st.caption("CSV / XLSX 파일 업로드 → BJ별 집계 엑셀 다운로드")

uploaded_files = st.file_uploader(
    "CSV 또는 XLSX 파일을 업로드하세요",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

def make_excel(df, bj_name):
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
        c = ws[f"{col}2"]
        c.border = border
        c.alignment = Alignment(horizontal="center")

    row = 3
    for _, r in df.iterrows():
        ws.cell(row=row, column=1, value=str(r["후원아이디"]))
        ws.cell(row=row, column=2, value=str(r["닉네임"]))

        heart = int(r["후원하트"])
        if heart < 0:
            heart = 0

        cell = ws.cell(row=row, column=3, value=heart)
        cell.number_format = "#,##0"
        row += 1

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 11

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

if uploaded_files:
    dfs = []
    for f in uploaded_files:
        try:
            if f.name.lower().endswith(".csv"):
                dfs.append(pd.read_csv(f))
            else:
                dfs.append(pd.read_excel(f))
        except Exception as e:
            st.error(f"{f.name} 읽기 실패: {e}")

    if dfs:
        merged = pd.concat(dfs, ignore_index=True)
        result = process_dataframe(merged)

        if not result:
            st.error("처리 결과가 없습니다.")
        else:
            st.success("집계 완료")

for bj, views in result.items():

    settlement_df = views["정산용"]
    bj_df = views["BJ용"]

    st.subheader(f"{bj}")

    st.download_button(
        label=f"{bj}_정산용.xlsx 다운로드",
        data=make_excel(settlement_df, bj),
        file_name=f"{bj}_정산용.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label=f"{bj}_BJ용.xlsx 다운로드",
        data=make_excel(bj_df, bj),
        file_name=f"{bj}_BJ용.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

