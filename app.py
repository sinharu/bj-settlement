import streamlit as st
import pandas as pd
from processor import process_dataframe
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment

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

# 업로드 안하면 여기서 멈춤 (에러 방지 핵심)
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
# 📊 웹 1차 요약표 (참여BJ별 하트 합산)
# ==================================================
try:
    tmp = merged.copy()

    col_idnick = next((c for c in tmp.columns if "후원" in c and "아이디" in c and "닉네임" in c), None)
    col_heart = next((c for c in tmp.columns if "후원" in c and "하트" in c), None)
    col_bj = next((c for c in tmp.columns if "참여" in c and "BJ" in c), None)

    if not (col_idnick and col_heart and col_bj):
        st.warning("필수 컬럼을 찾지 못했습니다.")
    else:
        tmp[col_heart] = pd.to_numeric(tmp[col_heart], errors="coerce").fillna(0)
        tmp.loc[tmp[col_heart] < 0, col_heart] = 0

        tmp["후원아이디"] = tmp[col_idnick].astype(str).str.replace(r"\(.*\)", "", regex=True).str.strip()

        def classify(x):
            s = str(x)
            if "@ka" in s:
                return "일반"
            if "@" in s:
                return "제휴"
            return "일반"

        tmp["구분"] = tmp["후원아이디"].apply(classify)

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
        pivot = pivot[["참여BJ", "일반", "제휴", "총합"]]
        pivot = pivot.sort_values("총합", ascending=False)

        # 숫자 포맷 보기 좋게
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
        ws[f"{col}2"].border = border
        ws[f"{col}2"].alignment = Alignment(horizontal="center")

    row = 3
    for _, r in df.iterrows():
        ws.cell(row=row, column=1, value=str(r["후원아이디"]))
        ws.cell(row=row, column=2, value=str(r["닉네임"]))
        heart = max(int(r["후원하트"]), 0)
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
