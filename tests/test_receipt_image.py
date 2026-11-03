import datetime as dt
from io import BytesIO

import pandas as pd
import pytest
from PIL import Image

import receipt_image as ri


def detail(*rows, dates=(dt.date(2026, 10, 23),), rounds=("15회차",)):
    df = pd.DataFrame(rows, columns=["닉네임", "후원하트", "후원종류"])
    df["방송일"] = [dates[i % len(dates)] for i in range(len(df))]
    df["방송회차"] = [rounds[i % len(rounds)] for i in range(len(df))]
    return df


def open_png(data):
    image = Image.open(BytesIO(data))
    image.load()
    return image


def test_summary_sums_by_nickname_and_breaks_ties_by_name():
    df = detail(
        ("아뇽", 261500, "투네이션"),
        ("다니엘", 100, "계좌후원"),
        ("냐옹", 261500, "투네이션"),
        ("다니엘", 900, "계좌후원"),
    )
    rows, _, _ = ri.receipt_summary(df)
    assert rows == [("냐옹", 261500), ("아뇽", 261500), ("다니엘", 1000)]


def test_summary_splits_account_from_every_other_type_and_skips_zero():
    df = detail(
        ("가", 1000, "투네이션"),
        ("나", 2000, "펀네이션"),
        ("다", 4000, " 계좌후원 "),
        ("라", 8000, "기타"),
        ("영원", 0, "투네이션"),
    )
    rows, web, account = ri.receipt_summary(df)
    assert (web, account) == (11000, 4000)
    assert "영원" not in [nick for nick, _ in rows]
    assert sum(amount for _, amount in rows) == web + account


@pytest.mark.parametrize(
    ("filename", "expected"),
    [
        ("10-16 위너컴퍼니 시즌2 월수금 11화_202610200114.xlsx", "11회차"),
        ("12회차 내역.xlsx", "12회차"),
        ("시즌2 내역.xlsx", None),
    ],
)
def test_round_from_filename(filename, expected):
    assert ri.round_from_filename(filename) == expected


def test_subtitle_single_broadcast():
    assert ri.receipt_subtitle(detail(("가", 1, "투네이션"))) == "2026-10-23 · 15회차"


def test_subtitle_multiple_broadcasts_shows_ranges():
    df = detail(
        ("가", 1, "투네이션"),
        ("나", 1, "투네이션"),
        dates=(dt.date(2026, 10, 23), dt.date(2026, 10, 26)),
        rounds=("9회차", "10회차"),
    )
    assert ri.receipt_subtitle(df) == "2026-10-23 ~ 2026-10-26 · 9~10회차"


def test_subtitle_without_round_is_date_only():
    df = detail(("가", 1, "투네이션"), rounds=(None,))
    assert ri.receipt_subtitle(df) == "2026-10-23"


def test_png_matches_sample_canvas_for_short_list():
    data = ri.make_receipt_png("다이아라운지", "희유", detail(("다니엘", 2254500, "계좌후원")))
    image = open_png(data)
    assert image.format == "PNG"
    assert image.size == (1080, 1920)


def test_many_rows_grow_canvas_and_keep_summary_box_below_table():
    rows = [(f"후원자{i}", 1000) for i in range(40)]
    image = open_png(ri.render_receipt_png("제목", "BJ", "2026-10-23", rows, 40000, 0))
    assert image.width == 1080
    table_bottom = ri.TABLE_LINE_Y + 3 + ri.ROW_HEIGHT * 40
    assert image.height == table_bottom + ri.MIN_TABLE_BOX_GAP + ri.BOX_HEIGHT + ri.BOX_BOTTOM_MARGIN


def test_long_nickname_never_reaches_amount_column():
    text, size = ri._fit_nickname("아주아주긴닉네임" * 12, 40, 500)
    assert ri._text_width(text, size) <= 500
    assert text.endswith("…")


def test_long_title_shrinks_to_fit_canvas():
    title = "다이아라운지 시즌2 월수금 후원 내역" * 2
    size = ri._fit_centered(title, 75)
    assert size < 75
    assert ri._text_width(title, size) <= ri.RIGHT_X - ri.MARGIN_X
    assert ri._fit_centered("다이아라운지", 75) == 75


def test_unique_filename_numbers_collisions():
    used = set()
    assert ri.unique_filename("A_B_내역이미지.png", used) == "A_B_내역이미지.png"
    assert ri.unique_filename("A_B_내역이미지.png", used) == "A_B_내역이미지_2.png"
    assert ri.unique_filename("A_B_내역이미지.png", used) == "A_B_내역이미지_3.png"


def test_unrenderable_nickname_falls_back_instead_of_tofu():
    # U+3FFFF는 어떤 글꼴에도 없는 코드포인트라 환경과 상관없이 같은 결과가 나온다.
    text, _ = ri._fit_nickname("\U0003ffff️", 40, 500)
    assert text == "?"
    mixed, _ = ri._fit_nickname("윤블리\U0003ffff", 40, 500)
    assert mixed == "윤블리"


def test_missing_korean_font_raises_clear_error(monkeypatch):
    monkeypatch.setattr(ri, "KOREAN_FONT_CANDIDATES", [("/없는/글꼴.ttf", 0)])
    monkeypatch.setattr(ri, "glob", lambda *a, **k: [])
    ri._font_chain.cache_clear()
    try:
        with pytest.raises(ri.ReceiptFontError):
            ri._font_chain()
    finally:
        ri._font_chain.cache_clear()
