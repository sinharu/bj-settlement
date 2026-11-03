"""BJ에게 전달하는 원화 후원 내역 이미지(PNG)를 만든다."""
from functools import lru_cache
from glob import glob
from io import BytesIO
from pathlib import Path
import re

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


ACCOUNT_TYPE = "계좌후원"

WIDTH = 1080
MIN_HEIGHT = 1920
MARGIN_X = 72
RIGHT_X = 1008
TEXT_RIGHT_X = 995
TABLE_LINE_Y = 473
ROW_HEIGHT = 73
BOX_HEIGHT = 283
BOX_BOTTOM_MARGIN = 150
MIN_TABLE_BOX_GAP = 60

BG = "#0e0d14"
BOX_FILL = "#1c1b22"
PINK = "#eac8ca"
SEPARATOR = "#27262c"
WHITE = "#ffffff"
GRAY = "#a5a5ab"

KOREAN_FONT_CANDIDATES = [
    # Streamlit Cloud (packages.txt의 fonts-noto-cjk)
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 1),
    ("/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc", 1),
    # 내 컴퓨터(Windows / macOS)
    ("C:/Windows/Fonts/malgunbd.ttf", 0),
    ("/System/Library/Fonts/AppleSDGothicNeo.ttc", 6),
]
SYMBOL_FONT_CANDIDATES = [
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 0),
    ("C:/Windows/Fonts/seguisym.ttf", 0),
]
IGNORED_CHARS = {"\ufe0f", "\ufe0e", "\u200d"}


class ReceiptFontError(RuntimeError):
    """한글을 그릴 글꼴을 찾지 못했을 때 발생한다."""


def _existing_fonts(candidates):
    return [(path, index) for path, index in candidates if Path(path).exists()]


@lru_cache(maxsize=1)
def _font_chain():
    korean = _existing_fonts(KOREAN_FONT_CANDIDATES)
    if not korean:
        found = sorted(glob("/usr/share/fonts/**/NotoSansCJK-Bold.ttc", recursive=True))
        korean = [(path, 1) for path in found]
    if not korean:
        raise ReceiptFontError("이미지에 쓸 한글 글꼴을 찾지 못했습니다.")
    return [korean[0], *_existing_fonts(SYMBOL_FONT_CANDIDATES)]


@lru_cache(maxsize=64)
def _font(path, index, size):
    return ImageFont.truetype(path, size=size, index=index)


@lru_cache(maxsize=4096)
def _font_for_char(char, size):
    """글자를 그릴 수 있는 첫 글꼴을 찾는다. 어떤 글꼴에도 없으면(이모지 등) None."""
    for path, index in _font_chain():
        font = _font(path, index, size)
        if bytes(font.getmask(char)) != bytes(font.getmask("\U0010ffff")):
            return font
    return None


def _runs(text, size):
    """(글꼴, 문자열) 묶음으로 나눈다. 어느 글꼴에도 없는 글자는 뺀다."""
    main = _font(*_font_chain()[0], size)
    runs = []
    for char in str(text):
        if char in IGNORED_CHARS:
            continue
        font = main if char.isspace() else _font_for_char(char, size)
        if font is None:
            continue
        if runs and runs[-1][0] is font:
            runs[-1][1] += char
        else:
            runs.append([font, char])
    return runs


def _text_width(text, size, spacing=0):
    width = 0.0
    count = 0
    for font, run in _runs(text, size):
        width += font.getlength(run)
        count += len(run)
    return width + spacing * max(count - 1, 0)


def _draw_text(draw, x, baseline, text, size, fill, align="left", spacing=0):
    width = _text_width(text, size, spacing)
    if align == "right":
        x -= width
    elif align == "center":
        x -= width / 2
    for font, run in _runs(text, size):
        pieces = run if spacing else [run]
        for piece in pieces:
            draw.text((x, baseline), piece, font=font, fill=fill, anchor="ls")
            x += font.getlength(piece) + spacing


def _fit_nickname(text, size, max_width, min_size=26):
    """닉네임이 금액 칸을 침범하지 않게 글자를 줄이고, 그래도 길면 말줄임한다."""
    drawable = "".join(run for _, run in _runs(text, size)).strip() or "?"
    while size > min_size and _text_width(drawable, size) > max_width:
        size -= 1
    if _text_width(drawable, size) <= max_width:
        return drawable, size
    while drawable and _text_width(drawable + "…", size) > max_width:
        drawable = drawable[:-1]
    return drawable + "…", size


def _fit_centered(text, size, min_size=24):
    """가운데 정렬 글자가 이미지 폭을 넘지 않을 크기를 찾는다."""
    while size > min_size and _text_width(text, size) > RIGHT_X - MARGIN_X:
        size -= 1
    return size


def _won(amount):
    return f"{int(amount):,}원"


def receipt_summary(detail_df: pd.DataFrame):
    """BJ 한 명의 표준 내역을 (닉네임별 합계 목록, 웹후원 합계, 계좌후원 합계)로 요약한다."""
    paid = detail_df[detail_df["후원하트"] > 0]
    nick_totals = paid.groupby("닉네임")["후원하트"].sum()
    rows = sorted(
        ((str(nick), int(total)) for nick, total in nick_totals.items()),
        key=lambda row: (-row[1], row[0]),
    )
    is_account = paid["후원종류"].astype(str).str.strip() == ACCOUNT_TYPE
    account_total = int(paid.loc[is_account, "후원하트"].sum())
    web_total = int(paid.loc[~is_account, "후원하트"].sum())
    return rows, web_total, account_total


def unique_filename(name, used: set) -> str:
    """이름이 달라도 파일명이 겹치면(A/B, A:B) ZIP에서 덮어써지므로 번호를 붙인다."""
    path = Path(name)
    unique, n = name, 2
    while unique in used:
        unique = f"{path.stem}_{n}{path.suffix}"
        n += 1
    used.add(unique)
    return unique


def round_from_filename(filename):
    """파일명의 '11화' 또는 '11회차'를 회차 이름으로 바꾼다."""
    match = re.search(r"(?<!\d)(\d+)\s*(?:회차|화)", Path(str(filename)).stem)
    return f"{int(match.group(1))}회차" if match else None


def _round_number(name):
    match = re.search(r"\d+", name)
    return int(match.group()) if match else None


def receipt_subtitle(detail_df: pd.DataFrame) -> str:
    """'2026-10-23 · 15회차' 형식. 여러 방송을 합치면 범위로 표시한다."""
    dates = sorted({d for d in detail_df["방송일"].dropna()})
    date_text = ""
    if dates:
        date_text = str(dates[0]) if len(dates) == 1 else f"{dates[0]} ~ {dates[-1]}"

    names = {str(name).strip() for name in detail_df["방송회차"].dropna() if str(name).strip()}
    numbers = sorted(_round_number(name) for name in names if _round_number(name) is not None)
    if len(names) <= 1:
        round_text = next(iter(names), "")
    elif len(numbers) == len(names):
        round_text = f"{numbers[0]}~{numbers[-1]}회차"
    else:
        round_text = ", ".join(sorted(names))
    return " · ".join(text for text in (date_text, round_text) if text)


def render_receipt_png(title, bj_name, subtitle, rows, web_total, account_total) -> bytes:
    table_top = TABLE_LINE_Y + 3
    table_bottom = table_top + ROW_HEIGHT * len(rows)
    height = max(MIN_HEIGHT, table_bottom + MIN_TABLE_BOX_GAP + BOX_HEIGHT + BOX_BOTTOM_MARGIN)
    box_top = height - BOX_BOTTOM_MARGIN - BOX_HEIGHT

    image = Image.new("RGB", (WIDTH, height), BG)
    draw = ImageDraw.Draw(image)

    center = WIDTH // 2
    _draw_text(draw, center, 197, title, _fit_centered(title, 75), WHITE, align="center")
    _draw_text(draw, center, 295, bj_name, _fit_centered(bj_name, 58), PINK, align="center")
    _draw_text(draw, center, 354, subtitle, _fit_centered(subtitle, 32), GRAY, align="center")

    _draw_text(draw, 143, 449, "후원자 닉네임", 29, GRAY, spacing=1.4)
    _draw_text(draw, TEXT_RIGHT_X + 3, 449, "누적 후원 금액", 29, GRAY, align="right", spacing=1.4)
    draw.rectangle([MARGIN_X, TABLE_LINE_Y, RIGHT_X - 1, TABLE_LINE_Y + 2], fill=PINK)

    for idx, (nickname, amount) in enumerate(rows):
        row_top = table_top + ROW_HEIGHT * idx
        baseline = row_top + 53
        amount_text = _won(amount)
        amount_width = _text_width(amount_text, 40)
        nick_text, nick_size = _fit_nickname(nickname, 40, TEXT_RIGHT_X - amount_width - 30 - 168)
        _draw_text(draw, 113, baseline, str(idx + 1), 38, GRAY, align="center")
        _draw_text(draw, 168, baseline, nick_text, nick_size, WHITE)
        _draw_text(draw, TEXT_RIGHT_X + 4, baseline, amount_text, 40, WHITE, align="right")
        sep_y = row_top + ROW_HEIGHT - 2
        draw.rectangle([MARGIN_X, sep_y, RIGHT_X - 1, sep_y + 1], fill=SEPARATOR)

    # 부드러운 둥근 테두리를 위해 크게 그린 뒤 줄인다.
    scale = 4
    box_w, box_h = RIGHT_X - MARGIN_X, BOX_HEIGHT
    box = Image.new("RGB", (box_w * scale, box_h * scale), BG)
    ImageDraw.Draw(box).rounded_rectangle(
        [0, 0, box_w * scale - 1, box_h * scale - 1],
        radius=30 * scale,
        fill=BOX_FILL,
        outline=PINK,
        width=3 * scale,
    )
    image.paste(box.resize((box_w, box_h), Image.LANCZOS), (MARGIN_X, box_top))

    _draw_text(draw, 112, box_top + 73, "총 후원 금액", 33, GRAY, spacing=0.8)
    _draw_text(draw, 112, box_top + 180, _won(web_total + account_total), 96, PINK)
    web_text = f"웹후원 {_won(web_total)}"
    _draw_text(draw, 112, box_top + 238, web_text, 32, GRAY)
    _draw_text(draw, 112 + _text_width(web_text, 32) + 26, box_top + 238, f"계좌후원 {_won(account_total)}", 32, GRAY)

    out = BytesIO()
    image.save(out, format="PNG", optimize=True)
    return out.getvalue()


def make_receipt_png(title, bj_name, detail_df: pd.DataFrame) -> bytes:
    rows, web_total, account_total = receipt_summary(detail_df)
    return render_receipt_png(title, bj_name, receipt_subtitle(detail_df), rows, web_total, account_total)
