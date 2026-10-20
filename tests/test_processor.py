from datetime import date, time
import unicodedata

import pandas as pd
import pytest
from pandas.testing import assert_frame_equal

from processor import (
    FORMAT_HEART,
    FORMAT_KRW,
    clean_and_prepare,
    detect_format,
    process_dataframe,
    to_standard,
)


STANDARD_COLUMNS = [
    "형식", "참여BJ", "회차", "후원시간", "날짜", "시간",
    "아이디", "닉네임", "후원하트", "구분", "후원종류",
]
VIEW_COLUMNS = ["아이디", "후원하트", "닉네임", "구분"]


@pytest.fixture
def heart_df():
    return pd.DataFrame({
        "후원시간": [
            "2026-10-19 18:00:00", "2026-10-19 18:01:00",
            "2026-10-19 18:02:00", "2026-10-19 18:03:00",
            "2026-10-19 18:04:00", "2026-10-19 18:05:00",
            "2026-10-19 18:06:00", "2026-10-19 18:07:00",
            "2026-10-19 18:08:00",
        ],
        "후원 아이디(닉네임)": [
            "abc@filejo(제휴팬)", "user(이전닉)", "user(대표닉)",
            "fan@ka(카카오팬)", "negative(음수)", "invalid(문자)",
            "other(다른팬)", "empty(빈BJ)", "missing(누락BJ)",
        ],
        "후원하트": ["500", "20", "70", "80", "-10", "문자", "40", 999, 999],
        "참여BJ": [" 가수 ", "가수", "가수", "가수", "가수", "가수", "댄서", " ", None],
        "하트점수": [0] * 9,
        "기여도": [0] * 9,
        "기타": [""] * 9,
        "업로드회차": ["2회차"] * 9,
    })


@pytest.fixture
def krw_df():
    return pd.DataFrame({
        "구분": list(range(1, 11)),
        "아이디": [
            "다른아이디", "팬", "팬", "계좌후원", "계좌후원",
            " 대체닉 ", None, "다른팬", "없음", "mail@filejo",
        ],
        "닉네임": [
            unicodedata.normalize("NFD", " 팬 "), "팬", "팬",
            "계좌후원", " 계좌후원 ", float("nan"), pd.NA,
            "다른팬", "제외팬", "mail@filejo",
        ],
        "채팅": [""] * 10,
        "스트리머": [
            " 가수 ", unicodedata.normalize("NFD", "가수"), "가수",
            "가수", "가수", "가수", "가수", " 댄서 ", " ", "가수",
        ],
        "종류": [
            " 투네이션 ", "펀네이션", "투네이션", "계좌후원", "계좌후원",
            "투네이션", "계좌후원", "펀네이션", "계좌후원", "투네이션",
        ],
        "개수": [1000, 2500, 99999, 500, 700, 300, 400, 600, 88888, 900],
        "기여도": [0] * 10,
        "배수": [1] * 10,
        "등록": [" 등록 ", "등록", "취소", "등록", "등록", "등록", "등록", "등록", "등록", "등록"],
        "메모": [""] * 10,
        "시간": ["2026-10-19 18:09:22.152"] * 10,
    })


def assert_results_equal(left, right):
    assert left.keys() == right.keys()
    for bj in left:
        assert left[bj].keys() == right[bj].keys()
        for view in left[bj]:
            assert_frame_equal(left[bj][view], right[bj][view])


def test_heart_detection_and_standardization(heart_df):
    before = heart_df.copy(deep=True)
    assert detect_format(heart_df) == FORMAT_HEART

    standard = to_standard(heart_df)

    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert standard.index.tolist() == list(range(7))
    assert standard["형식"].eq(FORMAT_HEART).all()
    assert standard["후원하트"].dtype == "int64"
    assert standard["후원하트"].tolist() == [500, 20, 70, 80, 0, 0, 40]
    assert standard["구분"].tolist() == ["제휴", "일반", "일반", "일반", "일반", "일반", "일반"]
    assert standard["후원종류"].eq("").all()
    assert standard["회차"].eq("2회차").all()
    assert standard["날짜"].eq(date(2026, 10, 19)).all()
    assert pd.api.types.is_datetime64_any_dtype(standard["후원시간"])
    assert_frame_equal(heart_df, before)


def test_heart_aggregation_representative_nickname_and_sorting(heart_df):
    result = process_dataframe(heart_df)

    assert set(result) == {"가수", "댄서"}
    settlement = result["가수"]["정산용"]
    bj_view = result["가수"]["BJ용"]
    assert settlement.columns.tolist() == VIEW_COLUMNS
    assert bj_view.columns.tolist() == VIEW_COLUMNS
    assert settlement["구분"].tolist() == ["일반", "일반", "일반", "일반", "제휴"]
    assert settlement["후원하트"].tolist() == [90, 80, 0, 0, 500]
    assert settlement.set_index("아이디").loc["user", "닉네임"] == "대표닉"
    assert bj_view["후원하트"].tolist() == [500, 90, 80, 0, 0]
    assert settlement["후원하트"].sum() == 670
    assert result["댄서"]["정산용"]["후원하트"].sum() == 40
    standard = to_standard(heart_df)
    for bj, views in result.items():
        assert_frame_equal(
            views["전체로그"],
            standard.loc[standard["참여BJ"] == bj].reset_index(drop=True),
        )


def test_heart_detection_keeps_substring_column_matching(heart_df):
    renamed = heart_df.rename(columns={
        "후원 아이디(닉네임)": "상세 후원 고객 아이디",
        "후원하트": "총 후원 하트 수량",
        "참여BJ": "참여 방송 BJ 이름",
        "후원시간": "기록 후원 시간",
    })
    assert detect_format(renamed) == FORMAT_HEART
    assert_frame_equal(to_standard(renamed), to_standard(heart_df))


def test_heart_missing_time_and_round_are_allowed(heart_df):
    standard = to_standard(heart_df.drop(columns=["후원시간", "업로드회차"]))

    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert pd.api.types.is_datetime64_any_dtype(standard["후원시간"])
    assert standard["후원시간"].isna().all()
    assert standard["날짜"].tolist() == [None] * len(standard)
    assert standard["시간"].tolist() == [None] * len(standard)
    assert standard["회차"].isna().all()


def test_krw_detection_and_standardization(krw_df):
    before = krw_df.copy(deep=True)
    assert detect_format(krw_df) == FORMAT_KRW

    standard = to_standard(krw_df)

    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert standard.index.tolist() == list(range(8))
    assert standard["형식"].eq(FORMAT_KRW).all()
    assert standard["후원하트"].dtype == "int64"
    assert standard["후원하트"].tolist() == [1000, 2500, 500, 700, 300, 400, 600, 900]
    assert standard["닉네임"].tolist() == [
        "팬", "팬", "계좌후원", "계좌후원", "대체닉", "", "다른팬", "mail@filejo",
    ]
    assert standard["아이디"].tolist() == standard["닉네임"].tolist()
    assert standard["구분"].eq("일반").all()
    assert standard["후원종류"].tolist() == [
        "투네이션", "펀네이션", "계좌후원", "계좌후원",
        "투네이션", "계좌후원", "펀네이션", "투네이션",
    ]
    assert standard["참여BJ"].tolist() == ["가수"] * 6 + ["댄서", "가수"]
    assert standard["회차"].isna().all()
    assert pd.api.types.is_datetime64_any_dtype(standard["후원시간"])
    assert standard["후원시간"].eq(pd.Timestamp("2026-10-19 18:09:22.152")).all()
    assert standard["날짜"].eq(date(2026, 10, 19)).all()
    assert standard["시간"].eq(time(18, 9, 22, 152000)).all()
    assert_frame_equal(krw_df, before)


def test_krw_aggregation_preserves_general_category_and_donation_types(krw_df):
    result = process_dataframe(krw_df)

    assert set(result) == {"가수", "댄서"}
    settlement = result["가수"]["정산용"]
    assert settlement.columns.tolist() == VIEW_COLUMNS
    assert settlement["아이디"].tolist() == ["팬", "계좌후원", "mail@filejo", "", "대체닉"]
    assert settlement["후원하트"].tolist() == [3500, 1200, 900, 400, 300]
    assert settlement["닉네임"].tolist() == settlement["아이디"].tolist()
    assert settlement["구분"].eq("일반").all()
    assert settlement["후원하트"].sum() == 6300
    assert result["댄서"]["정산용"]["후원하트"].sum() == 600
    assert_frame_equal(settlement, result["가수"]["BJ용"])
    standard = to_standard(krw_df)
    for bj, views in result.items():
        assert_frame_equal(
            views["전체로그"],
            standard.loc[standard["참여BJ"] == bj].reset_index(drop=True),
        )


def test_krw_strips_column_names(krw_df):
    padded = krw_df.rename(columns=lambda column: f" {column} ")
    assert detect_format(padded) == FORMAT_KRW
    assert_frame_equal(to_standard(padded), to_standard(krw_df))


def test_krw_optional_registration_and_id_columns(krw_df):
    raw = krw_df.drop(columns=["등록", "아이디"])
    standard = to_standard(raw)

    assert len(standard) == 9
    assert 99999 in standard["후원하트"].tolist()
    assert standard.loc[4:6, "닉네임"].tolist() == ["계좌후원", "", ""]


def test_krw_amount_cleanup_and_uploaded_round(krw_df):
    raw = krw_df.iloc[:2].copy()
    raw["개수"] = ["-500", "invalid"]
    raw["업로드회차"] = ["1회차", "2회차"]
    standard = to_standard(raw)

    assert standard["후원하트"].tolist() == [0, 0]
    assert standard["후원하트"].dtype == "int64"
    assert standard["회차"].tolist() == ["1회차", "2회차"]


@pytest.mark.parametrize("status", ["취소", "", None, "미등록"])
def test_krw_only_registered_rows_are_kept(krw_df, status):
    raw = krw_df.copy()
    raw["등록"] = status
    standard = to_standard(raw)

    assert standard.empty
    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert standard["후원하트"].dtype == "int64"
    assert process_dataframe(raw) is None


@pytest.mark.parametrize("columns", [
    [],
    ["아이디", "개수"],
    ["후원 아이디", "후원하트"],
    ["닉네임", "스트리머", "종류", "개수"],
    [0, 1],
])
def test_unrecognized_format_returns_none(columns):
    raw = pd.DataFrame(columns=columns)

    assert detect_format(raw) is None
    assert to_standard(raw) is None
    assert clean_and_prepare(raw) is None
    assert process_dataframe(raw) is None


@pytest.mark.parametrize("fixture_name", ["heart_df", "krw_df"])
def test_empty_recognized_input_returns_empty_standard_and_no_results(request, fixture_name):
    raw = request.getfixturevalue(fixture_name).iloc[:0]
    standard = to_standard(raw)

    assert standard.empty
    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert standard["후원하트"].dtype == "int64"
    assert process_dataframe(raw) is None


@pytest.mark.parametrize("fixture_name", ["heart_df", "krw_df"])
def test_standard_frame_is_copied_and_can_be_processed_again(request, fixture_name):
    raw = request.getfixturevalue(fixture_name)
    standard = to_standard(raw)
    before = standard.copy(deep=True)
    copied = clean_and_prepare(standard)

    assert copied is not standard
    assert_frame_equal(copied, standard)
    assert_results_equal(process_dataframe(raw), process_dataframe(standard))
    assert_frame_equal(standard, before)
    copied.loc[0, "후원하트"] = 123456789
    assert_frame_equal(standard, before)
