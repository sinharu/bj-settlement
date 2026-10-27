import re
from dataclasses import dataclass, field
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Border, Font, Side


TOTAL_LABEL = "총합계"
HEADER = ["아이디", "닉네임", "후원하트"]
HEART_NUMBER_FORMAT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
COLUMN_WIDTHS = {"A": 29.5, "B": 26.0, "C": 9.375}

_LINE_RE = re.compile(r"^(.*?)\s+(\d[\d,]*)\s*$")


@dataclass
class HeartText:
    name: str = ""
    name_total: int | None = None
    total: int | None = None
    rows: list[dict] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)

    @property
    def sum(self) -> int:
        return sum(r["후원하트"] for r in self.rows)


def _split_line(line: str):
    """한 줄을 "앞부분 + 마지막 숫자"로 나눈다."""
    m = _LINE_RE.match(line)
    if not m:
        return None
    return m.group(1).strip(), int(m.group(2).replace(",", ""))


def parse_heart_text(text: str) -> HeartText:
    """
    붙여넣은 후원 텍스트를 읽는다.
    - 첫 줄: 방송인 이름 + 전체 하트
    - 그 외: 아이디(닉네임) 하트
    - 총합계 줄은 합계 검증용
    하트 많은 순으로 정렬하고, 같으면 붙여넣은 순서를 유지한다.
    """
    result = HeartText()
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        parsed = _split_line(line)
        if parsed is None:
            result.skipped.append(line)
            continue
        head, value = parsed
        if re.sub(r"\s", "", head) == TOTAL_LABEL:
            result.total = value
            continue
        open_idx = head.find("(")
        close_idx = head.rfind(")")
        if open_idx > 0 and close_idx == len(head) - 1 and ")" not in head[:open_idx]:
            result.rows.append({
                "아이디": head[:open_idx].strip(),
                "닉네임": head[open_idx + 1:close_idx],
                "후원하트": value,
            })
        elif not result.name and not result.rows and "(" not in head and ")" not in head:
            result.name = head
            result.name_total = value
        else:
            result.skipped.append(line)

    result.rows.sort(key=lambda r: -r["후원하트"])  # sort()는 안정 정렬이라 동점은 입력 순서 유지
    return result


def _id_cell(user_id: str):
    # 숫자로만 된 아이디는 기존 엑셀처럼 숫자 셀로 넣는다 (앞자리 0이 사라지는 경우는 제외)
    return int(user_id) if re.fullmatch(r"0|[1-9]\d{0,14}", user_id) else user_id


def sheet_rows(result: HeartText) -> list[list]:
    top_total = result.name_total if result.name_total is not None else result.sum
    bottom_total = result.total if result.total is not None else result.sum
    return [
        [result.name, None, top_total],
        list(HEADER),
        *[[_id_cell(r["아이디"]), r["닉네임"], r["후원하트"]] for r in result.rows],
        [TOTAL_LABEL, None, bottom_total],
    ]


def total_mismatch(result: HeartText) -> int | None:
    """적힌 합계(총합계 줄, 없으면 첫 줄)가 실제 합과 다르면 적힌 값을 돌려준다."""
    expected = result.total if result.total is not None else result.name_total
    if expected is not None and expected != result.sum:
        return expected
    return None


def make_heart_excel(result: HeartText) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    font = Font(name="맑은 고딕", size=11)

    for r_idx, row in enumerate(sheet_rows(result), start=1):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if isinstance(value, str):
                cell.data_type = "s"  # "="로 시작하는 닉네임이 수식으로 저장되지 않게
            cell.border = border
            cell.font = font
            if c_idx == 3:
                cell.number_format = HEART_NUMBER_FORMAT

    for col, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
