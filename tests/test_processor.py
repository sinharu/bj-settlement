from datetime import date, datetime, time
from io import BytesIO
import unicodedata

import pandas as pd
import pytest
from openpyxl import Workbook
from pandas.testing import assert_frame_equal

from processor import (
    FORMAT_HEART,
    FORMAT_KRW,
    broadcast_start_date,
    clean_and_prepare,
    detect_format,
    export_time_from_filename,
    process_dataframe,
    to_standard,
    written_date_from_filename,
)


STANDARD_COLUMNS = [
    "형식", "참여BJ", "회차", "후원시간", "날짜", "시간",
    "아이디", "닉네임", "후원하트", "구분", "후원종류",
    "방송일", "방송회차",
]
VIEW_COLUMNS = ["아이디", "후원하트", "닉네임", "구분"]


def _read_ledger_xlsx(
    *, title="희유 총 후원 내역", broadcast_date="2026-10-23", broadcast_round="15회차",
    rows=None, footer_rows=None, header=None,
):
    """고객 파일 없이 첫 행을 열 이름으로 읽는 앱의 XLSX 입력을 재현한다."""
    if rows is None:
        rows = [
            (1, datetime(2026, 10, 24, 15, 1, 0, 125000), "계좌후원", "계좌팬",
             "계좌후원", 0, 119000, 119000, "synthetic:bank"),
            (4, datetime(2026, 10, 24, 15, 2), "투네이션", "웹팬",
             "웹후원", 17000, 0, 17000, "synthetic:web"),
        ]
    if footer_rows is None:
        footer_rows = [("웹후원 합계", 17000), ("계좌후원 합계", 119000), ("총 후원 합계", 136000)]
    if header is None:
        header = [
            "번호", "후원 일시", "플랫폼", "후원자 닉네임", "후원 구분",
            "웹후원 금액", "계좌후원 금액", "총 후원 금액", "원본 이벤트 ID",
        ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "총 후원 내역"
    for row in [
        [title], ["방송 회차명", broadcast_round], ["방송일", broadcast_date],
        ["생성 시각", datetime(2026, 10, 24, 16)], ["ledgerVersion", 20597],
        ["집계 기준", "후원 확정 시점 원화(정수)"], ["포함 상태", "확정·미배정·수동"],
        header, *rows, [], *footer_rows,
    ]:
        sheet.append(row)
    workbook.create_sheet("읽지 않는 시트").append(["이 시트의 데이터는 집계하지 않음"])
    with BytesIO() as stream:
        workbook.save(stream)
        workbook.close()
        stream.seek(0)
        return pd.read_excel(stream)


@pytest.fixture
def ledger_df():
    return _read_ledger_xlsx()


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
    assert standard["방송일"].tolist() == [None] * len(standard)
    assert standard["방송회차"].tolist() == [None] * len(standard)
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
    assert standard["방송일"].tolist() == [None] * len(standard)
    assert standard["방송회차"].tolist() == [None] * len(standard)
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


@pytest.mark.parametrize("fixture_name", ["heart_df", "krw_df", "ledger_df"])
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


def test_ledger_detection_and_standardization(ledger_df):
    before = ledger_df.copy(deep=True)
    assert ledger_df.columns.tolist() == ["희유 총 후원 내역", *[f"Unnamed: {i}" for i in range(1, 9)]]
    assert detect_format(ledger_df) == FORMAT_KRW

    standard = to_standard(ledger_df)

    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert standard.index.tolist() == [0, 1]
    assert standard["형식"].tolist() == [FORMAT_KRW, FORMAT_KRW]
    assert standard["참여BJ"].tolist() == ["희유", "희유"]
    assert standard["닉네임"].tolist() == ["계좌팬", "웹팬"]
    assert standard["아이디"].tolist() == standard["닉네임"].tolist()
    assert standard["후원하트"].tolist() == [119000, 17000]
    assert standard["후원하트"].dtype == "int64"
    assert standard["후원종류"].tolist() == ["계좌후원", "투네이션"]
    assert standard["구분"].tolist() == ["일반", "일반"]
    assert standard["회차"].tolist() == [None, None]
    assert standard["방송일"].tolist() == [date(2026, 10, 23)] * 2
    assert all(type(value) is date for value in standard["방송일"])
    assert standard["방송회차"].tolist() == ["15회차"] * 2
    assert standard["후원시간"].tolist() == [
        pd.Timestamp("2026-10-24 15:01:00.125"), pd.Timestamp("2026-10-24 15:02:00"),
    ]
    assert standard["날짜"].tolist() == [date(2026, 10, 24)] * 2
    assert standard["시간"].tolist() == [time(15, 1, 0, 125000), time(15, 2)]
    assert_frame_equal(ledger_df, before)

    result = process_dataframe(ledger_df)
    assert list(result) == ["희유"]
    assert result["희유"]["정산용"]["후원하트"].sum() == 136000
    assert_frame_equal(result["희유"]["전체로그"], standard)


@pytest.mark.parametrize("title", [None, 123, "총 후원 내역", "  총 후원 내역  ", "희유 총 후원 내역 메모"])
def test_ledger_detection_requires_title_with_bj_name(title):
    raw = _read_ledger_xlsx(title=title)
    assert detect_format(raw) is None
    assert to_standard(raw) is None


@pytest.mark.parametrize("missing_header", ["후원 일시", "플랫폼", "후원자 닉네임", "총 후원 금액"])
def test_ledger_detection_requires_all_core_headers_on_one_row(ledger_df, missing_header):
    raw = ledger_df.copy()
    raw.iloc[6] = raw.iloc[6].replace(missing_header, "다른 열")
    raw.iloc[0, 0] = missing_header
    assert detect_format(raw) is None


def test_ledger_detection_accepts_the_four_minimum_headers():
    raw = _read_ledger_xlsx(
        header=["후원 일시", "플랫폼", "후원자 닉네임", "총 후원 금액"],
        rows=[], footer_rows=[],
    )
    assert detect_format(raw) == FORMAT_KRW


def test_ledger_title_and_headers_are_trimmed_and_index_is_positional(ledger_df):
    raw = ledger_df.copy()
    raw.columns = [unicodedata.normalize("NFD", "  희유") + " 총 후원 내역  ", *raw.columns[1:]]
    raw.iloc[6] = raw.iloc[6].map(lambda value: f" {value} ")
    raw.index = range(100, 100 + len(raw))
    assert detect_format(raw) == FORMAT_KRW
    assert_frame_equal(to_standard(raw), to_standard(ledger_df))


@pytest.mark.parametrize("missing_header", ["번호", "웹후원 금액", "계좌후원 금액"])
def test_ledger_missing_amount_or_number_column_raises_value_error(ledger_df, missing_header):
    raw = ledger_df.copy()
    raw.iloc[6] = raw.iloc[6].replace(missing_header, "다른 열")
    assert detect_format(raw) == FORMAT_KRW
    with pytest.raises(ValueError, match=f"필수 열.*{missing_header}"):
        to_standard(raw)


@pytest.mark.parametrize("platform, web_type", [
    (" 투네이션 ", "투네이션"), ("다른플랫폼", "다른플랫폼"),
    (None, "웹후원"), ("  ", "웹후원"), (" 계좌후원 ", "웹후원"),
])
def test_ledger_splits_web_and_bank_amounts_and_preserves_platform(platform, web_type):
    raw = _read_ledger_xlsx(rows=[
        (3, datetime(2026, 10, 24, 15, 1), platform, "합성팬@filejo",
         "참고용 구분", 17000, 119000, 136000, "synthetic:mixed"),
    ])
    standard = to_standard(raw)

    assert standard["후원하트"].tolist() == [17000, 119000]
    assert standard["후원종류"].tolist() == [web_type, "계좌후원"]
    assert standard["닉네임"].tolist() == ["합성팬@filejo"] * 2
    assert standard["구분"].tolist() == ["일반"] * 2
    assert standard["후원시간"].tolist() == [pd.Timestamp("2026-10-24 15:01:00")] * 2
    assert standard["방송일"].tolist() == [date(2026, 10, 23)] * 2
    assert standard["방송회차"].tolist() == ["15회차"] * 2
    assert process_dataframe(raw)["희유"]["정산용"]["후원하트"].tolist() == [136000]


def test_ledger_only_numeric_number_rows_are_donations():
    raw = _read_ledger_xlsx(
        rows=[
            ("7", datetime(2026, 10, 24, 15), "투네이션", "숫자번호", "웹후원", 42, 0, 42, "synthetic:7"),
            ("참고", None, "투네이션", "참고닉", None, "금액 아님", None, None),
            (None, None, "투네이션", "번호없음", None, 500, 0, 500),
        ],
        footer_rows=[("웹후원 합계", 42), ("계좌후원 합계", 0), ("총 후원 합계", 42)],
    )
    standard = to_standard(raw)
    assert standard["닉네임"].tolist() == ["숫자번호"]
    assert standard["후원하트"].tolist() == [42]


@pytest.mark.parametrize("label, expected", [
    ("웹후원 합계", 17000), ("계좌후원 합계", 119000), ("총 후원 합계", 136000),
])
@pytest.mark.parametrize("delta", [-1, 1000])
def test_ledger_each_present_footer_total_is_validated(label, expected, delta):
    raw = _read_ledger_xlsx(footer_rows=[(label, expected + delta)])
    with pytest.raises(ValueError) as error:
        process_dataframe(raw)
    message = str(error.value)
    assert label in message
    assert f"하단 합계 {expected + delta:,}원" in message
    assert f"데이터 합계 {expected:,}원" in message
    assert f"차이 {abs(delta):,}원" in message


@pytest.mark.parametrize("web, bank, total", [(17001, 119000, 136000), (17000, 119001, 136000), (17000, 119000, 136001)])
def test_ledger_row_mismatch_raises_even_without_footer(web, bank, total):
    raw = _read_ledger_xlsx(
        rows=[(3, datetime(2026, 10, 24, 15), "투네이션", "합성팬", "웹후원", web, bank, total, "synthetic:bad")],
        footer_rows=[],
    )
    with pytest.raises(ValueError) as error:
        to_standard(raw)
    message = str(error.value)
    assert "엑셀 9행" in message
    assert f"웹후원 금액 {web:,}원" in message
    assert f"계좌후원 금액 {bank:,}원" in message
    assert f"총 후원 금액 {total:,}원" in message
    assert "차이 1원" in message


@pytest.mark.parametrize("amount", [None, "금액오류", -1, 0.5, "Infinity"])
def test_ledger_invalid_amounts_are_not_silently_coerced(ledger_df, amount):
    raw = ledger_df.copy()
    raw.iloc[7, 5] = amount
    with pytest.raises(ValueError, match="엑셀 9행 웹후원 금액"):
        to_standard(raw)


@pytest.mark.parametrize("amount", [None, "금액오류", -1, 136000.5])
def test_ledger_invalid_footer_amount_is_not_treated_as_missing_footer(amount):
    raw = _read_ledger_xlsx(footer_rows=[("총 후원 합계", amount)])
    with pytest.raises(ValueError, match="총 후원 합계.*정수"):
        to_standard(raw)


def test_ledger_missing_footer_totals_are_allowed(ledger_df):
    raw = _read_ledger_xlsx(footer_rows=[])
    assert_frame_equal(to_standard(raw), to_standard(ledger_df))


def test_ledger_empty_donations_with_zero_footer_totals():
    raw = _read_ledger_xlsx(
        rows=[], footer_rows=[("웹후원 합계", 0), ("계좌후원 합계", 0), ("총 후원 합계", 0)],
    )
    assert detect_format(raw) == FORMAT_KRW
    standard = to_standard(raw)
    assert standard.empty
    assert standard.columns.tolist() == STANDARD_COLUMNS
    assert standard["후원하트"].dtype == "int64"
    assert process_dataframe(raw) is None


@pytest.mark.parametrize("value, expected", [
    (None, None), ("잘못된 날짜", None), (datetime(2026, 10, 23, 20), date(2026, 10, 23)),
])
def test_ledger_broadcast_date_parsing(value, expected):
    standard = to_standard(_read_ledger_xlsx(broadcast_date=value))
    assert standard["방송일"].tolist() == [expected] * 2


@pytest.mark.parametrize("value, expected", [(None, None), (" 15회차 ", "15회차"), (15, "15")])
def test_ledger_broadcast_round_is_text(value, expected):
    standard = to_standard(_read_ledger_xlsx(broadcast_round=value))
    assert standard["방송회차"].tolist() == [expected] * 2


def test_ledger_uploaded_round_is_separate_from_broadcast_round(ledger_df):
    raw = ledger_df.copy()
    raw["업로드회차"] = "업로드 1회차"
    standard = to_standard(raw)
    assert standard["회차"].tolist() == ["업로드 1회차"] * 2
    assert standard["방송회차"].tolist() == ["15회차"] * 2


@pytest.mark.parametrize("fixture_name", ["heart_df", "krw_df"])
def test_legacy_standard_frame_without_broadcast_metadata_is_still_supported(request, fixture_name):
    raw = request.getfixturevalue(fixture_name)
    legacy = to_standard(raw).drop(columns=["방송일", "방송회차"])
    before = legacy.copy(deep=True)
    assert_results_equal(process_dataframe(raw), process_dataframe(legacy))
    assert_frame_equal(legacy, before)


@pytest.mark.parametrize(
    ("stem", "expected"),
    [
        ("비너스 1회차 26.11.03_202611040505", date(2026, 11, 3)),
        ("26.11.03 비너스", date(2026, 11, 3)),
        ("26-11-03 비너스", date(2026, 11, 3)),
        ("2026.11.03 비너스", date(2026, 11, 3)),
        ("비너스 1회차 2026-11-03_202611040505", date(2026, 11, 3)),
        ("11.03 비너스", date(2026, 11, 3)),
        ("11-03 비너스", date(2026, 11, 3)),
        ("비너스 1회차 11.3", date(2026, 11, 3)),
        ("11월 3일 내역", date(2026, 11, 3)),
        ("시즌2 11.03 내역", date(2026, 11, 3)),
        ("10-16 위너컴퍼니 시즌2 월수금 11화_202610200114", date(2026, 10, 16)),
        ("10-08 7화_202610090235", date(2026, 10, 8)),
        ("10-08-7화", date(2026, 10, 8)),
        ("11.10 후니_내역", date(2026, 11, 10)),
        ("11.03.26 내역", date(2026, 11, 3)),
        ("26.13.45 11.03 내역", date(2026, 11, 3)),
        ("99.99 내역_202611040505", None),
        ("시즌2 내역", None),
    ],
)
def test_written_date_from_filename(stem, expected):
    assert written_date_from_filename(stem, this_year=2026) == expected


def test_export_time_from_filename():
    assert export_time_from_filename("비너스_202611040505") == pd.Timestamp(2026, 11, 4, 5)
    assert export_time_from_filename("비너스_20261399") is None
    assert export_time_from_filename("비너스") is None


@pytest.mark.parametrize(
    ("times", "expected"),
    [
        # 저녁에 시작해 다음 날 새벽에 끝난 방송
        (["2026-11-03 18:00", "2026-11-04 05:00"], date(2026, 11, 3)),
        # 후원이 자정 이후에만 있는 경우
        (["2026-11-04 00:30", "2026-11-04 05:00"], date(2026, 11, 3)),
        # 14시 전에 시작해도 전날로 밀리지 않는다
        (["2026-11-03 13:58", "2026-11-03 15:00", "2026-11-04 01:00"], date(2026, 11, 3)),
        # 다음 날 14시를 넘기고 후원이 더 많아도 시작한 날짜다
        (["2026-11-03 15:00", "2026-11-04 15:00", "2026-11-04 16:00", "2026-11-04 17:00"], date(2026, 11, 3)),
    ],
)
def test_broadcast_start_date(times, expected):
    series = pd.Series(pd.to_datetime(times))
    business = series.apply(lambda t: (t - pd.Timedelta(days=1) if t.hour < 14 else t).date())
    assert broadcast_start_date(series, business) == expected
