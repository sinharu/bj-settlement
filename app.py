import streamlit as st
import pandas as pd
from processor import process_dataframe
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment

st.set_page_config(page_title="BJ 하트 집계", layout="centered")

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

    # 1행
    ws["A1"] = ""
    ws["B1"] = bj_name
    ws["C1"] = total

    # 2행 헤더
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
        c = ws.cell(row=row, column=3, value=heart)
        c.number_format = "#,##0"

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

            for bj, out_df in result.items():
                st.download_button(
                    label=f"{bj}.xlsx 다운로드",
                    data=make_excel(out_df, bj),
                    file_name=f"{bj}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

