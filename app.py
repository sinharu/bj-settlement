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
# 📦 엑셀 공통 유틸 (콤마/테두리/열너비)
# ==================================================
thin = Side(style="thin")
all_border = Border(left=thin, right=thin, top=thin, bottom=thin)

def apply_border(ws):
    # "값이 있는 셀" 전부 테두리
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value not in (None, ""):
                cell.border = all_border

def auto_width(ws, min_w=18, max_w=45, pad=4):
    # 기본 넓이 유지 + 데이터 길이 따라 자동 확장
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value not in (None, ""):
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + pad, min_w), max_w)

def format_header_row(ws, header_row=1):
    # 헤더 정렬(선택)
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=c)
        if cell.value not in (None, ""):
            cell.alignment = Alignment(horizontal="center")


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


# ==================================================
# 📅 파일 1개 업로드 시 날짜 prefix (파일명 우선 → 없으면 데이터 최솟날짜)
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
    prefix = None  # 여러개면 prefix 안 붙임


# ==================================================
# 🧠 하트 구분 / 아이디-닉네임 분리
# ==================================================
def classify_heart_type(user_id: str) -> str:
    s = str(user_id)
    if "@ka" in s:
        return "일반"
    if "@" in s:
        return "제휴"
    return "일반"

def split_id_nickname(text):
    text = str(text)
    if "(" in text and ")" in text:
        id_part, nick_part = text.split("(", 1)
        nick_part = nick_part.rstrip(")")
    else:
        id_part = text
        nick_part = ""
    return id_part.strip(), nick_part.strip()


# ==================================================
# 📊 웹 요약표 (참여BJ별 일반/제휴/총합)
# ==================================================
try:
    tmp = merged.copy()

    col_id = next((c for c in tmp.columns if "후원" in c and "아이디" in c), None)
    col_heart = next((c for c in tmp.columns if "후원" in c and "하트" in c), None)
    col_bj = next((c for c in tmp.columns if "참여" in c and "BJ" in c), None)

    if not (col_id and col_heart and col_bj):
        st.warning("요약표: 필수 컬럼(후원아이디/후원하트/참여BJ)을 찾지 못했습니다.")
    else:
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

        # 화면용 콤마(문자열) — 엑셀은 number_format으로 처리하니까 여기만 문자열로 OK
        for c in ["일반", "제휴", "총합"]:
            pivot[c] = pivot[c].apply(lambda x: f"{int(x):,}")

        st.subheader("요약_참여BJ_총계")
        st.dataframe(pivot.reset_index(drop=True), hide_index=True, use_container_width=True)

except Exception as e:
    st.warning(f"요약표 생성 중 오류: {e}")


# ==================================================
# 📁 BJ별 파일 생성 (정산용 / BJ용) - 콤마/테두리/열너비 적용
# ==================================================
result = process_dataframe(merged)

if not result:
    st.error("집계 결과가 없습니다.")
    st.stop()

def make_excel(df: pd.DataFrame, bj_name: str) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "정산표"

    # 상단 합계
    total = int(pd.to_numeric(df["후원하트"], errors="coerce").fillna(0).sum())
    ws["A1"] = ""
    ws["B1"] = bj_name
    total_cell = ws["C1"]
    total_cell.value = total
    total_cell.number_format = "#,##0"

    # 헤더
    ws.append(["후원아이디", "닉네임", "후원하트"])
    format_header_row(ws, header_row=2)

    # 데이터
    for _, r in df.iterrows():
        row = ws.max_row + 1
        ws.cell(row=row, column=1, value=str(r.get("후원아이디", "")))
        ws.cell(row=row, column=2, value=str(r.get("닉네임", "")))

        heart = pd.to_numeric(r.get("후원하트", 0), errors="coerce")
        heart = 0 if pd.isna(heart) else int(heart)
        heart = max(heart, 0)

        cell = ws.cell(row=row, column=3, value=heart)
        cell.number_format = "#,##0"

    # 기본 폭(너무 좁아지는 것 방지) + 자동 보정
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 16
    auto_width(ws, min_w=18, max_w=45, pad=4)
    apply_border(ws)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ==================================================
# 📦 총합산 파일 (여러 파일 업로드 시) - 3시트 구조
# 1) 일자별집계  2) 총합  3) BJ별 상세(각 BJ 1시트)
# ==================================================
def make_total_excel(df: pd.DataFrame) -> BytesIO | None:
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
    tmp.loc[tmp[col_heart] < 0, col_heart] = 0

    tmp[["아이디", "닉네임"]] = tmp[col_idnick].apply(lambda x: pd.Series(split_id_nickname(x)))
    tmp["구분"] = tmp["아이디"].apply(classify_heart_type)

    # 1) 일자별집계
    ws1 = wb.create_sheet("일자별집계")
    ws1.append(["날짜", "BJ", "일반", "제휴", "총합"])
    format_header_row(ws1, 1)

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
        row = ws1.max_row + 1
        ws1.cell(row=row, column=1, value=r["날짜"])
        ws1.cell(row=row, column=2, value=r[col_bj])

        c1 = ws1.cell(row=row, column=3, value=int(r["일반"]))
        c2 = ws1.cell(row=row, column=4, value=int(r["제휴"]))
        c3 = ws1.cell(row=row, column=5, value=int(r["총합"]))
        for c in (c1, c2, c3):
            c.number_format = "#,##0"

    ws1.column_dimensions["A"].width = 20
    ws1.column_dimensions["B"].width = 28
    ws1.column_dimensions["C"].width = 16
    ws1.column_dimensions["D"].width = 16
    ws1.column_dimensions["E"].width = 16
    auto_width(ws1, min_w=18, max_w=45, pad=4)
    apply_border(ws1)

    # 2) 총합
    ws2 = wb.create_sheet("총합")
    ws2.append(["BJ", "일반", "제휴", "총합"])
    format_header_row(ws2, 1)

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
        row = ws2.max_row + 1
        ws2.cell(row=row, column=1, value=r[col_bj])

        c1 = ws2.cell(row=row, column=2, value=int(r["일반"]))
        c2 = ws2.cell(row=row, column=3, value=int(r["제휴"]))
        c3 = ws2.cell(row=row, column=4, value=int(r["총합"]))
        for c in (c1, c2, c3):
            c.number_format = "#,##0"

    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 16
    ws2.column_dimensions["D"].width = 16
    auto_width(ws2, min_w=18, max_w=45, pad=4)
    apply_border(ws2)

    # 3) BJ별 상세 (각 BJ 1시트)
    for bj in tmp[col_bj].dropna().unique():
        ws = wb.create_sheet(str(bj))
        sub = tmp[tmp[col_bj] == bj].copy()

        normal_sum = int(sub[sub["구분"] == "일반"][col_heart].sum())
        partner_sum = int(sub[sub["구분"] == "제휴"][col_heart].sum())
        total_sum = normal_sum + partner_sum

        # 상단 한 줄 표시(일렬)
        ws["A1"] = "총하트"
        ws["B1"] = total_sum
        ws["C1"] = "일반하트"
        ws["D1"] = normal_sum
        ws["E1"] = "제휴하트"
        ws["F1"] = partner_sum
        ws["B1"].number_format = "#,##0"
        ws["D1"].number_format = "#,##0"
        ws["F1"].number_format = "#,##0"

        ws.append([])
        ws.append(["날짜", "시간", "아이디", "닉네임", "하트", "구분"])
        format_header_row(ws, header_row=3)

        # 정렬(원하면 여기서 날짜/시간 정렬)
        sub = sub.sort_values(by=[col_time], ascending=True)

        for _, r in sub.iterrows():
            row = ws.max_row + 1
            ws.cell(row=row, column=1, value=r["날짜"])
            ws.cell(row=row, column=2, value=r["시간"])
            ws.cell(row=row, column=3, value=r["아이디"])
            ws.cell(row=row, column=4, value=r["닉네임"])

            h = int(r[col_heart])
            hc = ws.cell(row=row, column=5, value=h)
            hc.number_format = "#,##0"

            ws.cell(row=row, column=6, value=r["구분"])

        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 32
        ws.column_dimensions["D"].width = 26
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 12
        auto_width(ws, min_w=18, max_w=45, pad=4)
        apply_border(ws)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ==================================================
# 📥 다운로드 UI
# ==================================================
st.success("집계 완료")

# 여러 파일 업로드일 때만 총합산 제공(요구사항)
if len(uploaded_files) > 1:
    total_file = make_total_excel(merged)
    if total_file is None:
        st.warning("총합산 생성 실패: 필수 컬럼(후원시간/후원아이디/후원하트/참여BJ)을 찾지 못했습니다.")
    else:
        st.download_button(
            label="총합산.xlsx 다운로드",
            data=total_file,
            file_name="총합산.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# BJ별 파일 제공 (파일 1개일 때만 prefix 붙임)
for bj, views in result.items():
    st.subheader(bj)

    filename1 = f"{prefix}_{bj}_정산용.xlsx" if prefix else f"{bj}_정산용.xlsx"
    filename2 = f"{prefix}_{bj}_BJ용.xlsx" if prefix else f"{bj}_BJ용.xlsx"

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
