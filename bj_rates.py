import math

import pandas as pd


SETTLE_COL = "정산비율 (%)"
SUPPORT_COL = "협력지원율 (%)"
RATE_COLUMNS = {SETTLE_COL: "정산비율", SUPPORT_COL: "협력지원율"}


def build_rate_table(bj_names, settle_pct: float, support_pct: float) -> pd.DataFrame:
    """BJ별 비율 입력표의 초기값. 모든 BJ를 기본 비율(%)로 채운다."""
    names = list(bj_names)
    return pd.DataFrame({
        "BJ": names,
        SETTLE_COL: [float(settle_pct)] * len(names),
        SUPPORT_COL: [float(support_pct)] * len(names),
    })


def _pct_to_rate(value, fallback: float) -> float:
    """% 입력값을 비율로 바꾼다. 비었거나 숫자가 아니면 기본값, 0~100 밖은 잘라낸다."""
    number = pd.to_numeric(value, errors="coerce")
    if number is None or (isinstance(number, float) and math.isnan(number)):
        return fallback
    return min(max(float(number), 0.0), 100.0) / 100


def rates_by_bj(table: pd.DataFrame, base_rates: dict) -> dict[str, dict]:
    """
    입력표를 BJ별 비율 dict로 바꾼다.
    정산비율·협력지원율만 BJ마다 다르고, 수수료·부가세 등 나머지는 base_rates를 그대로 쓴다.
    """
    result = {}
    for row in table.to_dict("records"):
        rates = dict(base_rates)
        for col, key in RATE_COLUMNS.items():
            rates[key] = _pct_to_rate(row.get(col), base_rates[key])
        result[str(row["BJ"])] = rates
    return result
