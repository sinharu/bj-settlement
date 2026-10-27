import pandas as pd
import pytest
from pandas.testing import assert_frame_equal

from bj_rates import build_rate_table, rates_by_bj


SETTLE_COL = "정산비율 (%)"
SUPPORT_COL = "협력지원율 (%)"
TABLE_COLUMNS = ["BJ", SETTLE_COL, SUPPORT_COL]


@pytest.fixture
def base_rates():
    return {
        "정산비율": 0.45,
        "협력지원율": 0.07,
        "투네이션수수료": 0.035,
        "펀네이션수수료": 0.05,
        "계좌후원부가세": 0.10,
    }


@pytest.mark.parametrize(("settle_pct", "support_pct"), [
    (45, 7),
    (37.5, 12.25),
])
def test_build_rate_table_preserves_names_order_and_float_defaults(
    settle_pct, support_pct,
):
    names = ["BJ나", "BJ가", "BJ다"]

    table = build_rate_table(names, settle_pct, support_pct)

    assert table.columns.tolist() == TABLE_COLUMNS
    assert_frame_equal(table, pd.DataFrame({
        "BJ": ["BJ나", "BJ가", "BJ다"],
        SETTLE_COL: [float(settle_pct)] * 3,
        SUPPORT_COL: [float(support_pct)] * 3,
    }))
    assert names == ["BJ나", "BJ가", "BJ다"]


def test_empty_names_produce_empty_table_and_rate_mapping(base_rates):
    before = base_rates.copy()

    table = build_rate_table([], 45, 7)

    assert table.empty
    assert table.columns.tolist() == TABLE_COLUMNS
    assert rates_by_bj(table, base_rates) == {}
    assert base_rates == before


def test_initial_percentages_override_base_rates_and_preserve_other_keys(base_rates):
    table = build_rate_table(["BJ가", "BJ나"], 60, 15)

    result = rates_by_bj(table, base_rates)

    assert set(result) == {"BJ가", "BJ나"}
    assert all(isinstance(name, str) for name in result)
    for rates in result.values():
        assert rates == {**base_rates, "정산비율": 0.60, "협력지원율": 0.15}
        assert isinstance(rates["정산비율"], float)
        assert isinstance(rates["협력지원율"], float)


def test_edited_cells_apply_to_their_bj_even_with_reordered_rows(base_rates):
    table = build_rate_table(["BJ가", "BJ나", "BJ다"], 45, 7)
    table.loc[0, SETTLE_COL] = 37.5
    table.loc[0, SUPPORT_COL] = 12.25
    table.loc[1, SETTLE_COL] = 62.5
    table.loc[1, SUPPORT_COL] = 0
    edited = table.loc[[2, 0, 1]].copy()

    result = rates_by_bj(edited, base_rates)

    assert result == {
        "BJ가": {**base_rates, "정산비율": 0.375, "협력지원율": 0.1225},
        "BJ나": {**base_rates, "정산비율": 0.625, "협력지원율": 0.0},
        "BJ다": base_rates,
    }


@pytest.mark.parametrize(("column", "key"), [
    (SETTLE_COL, "정산비율"),
    (SUPPORT_COL, "협력지원율"),
])
@pytest.mark.parametrize(("value", "expected"), [
    (0, 0.0),
    (100, 1.0),
    (37.5, 0.375),
    (12.34, 0.1234),
    ("50", 0.5),
    (" 37.5 ", 0.375),
    (-0.01, 0.0),
    (100.01, 1.0),
    ("-20", 0.0),
    ("125", 1.0),
    (float("-inf"), 0.0),
    (float("inf"), 1.0),
])
def test_numeric_edits_convert_percentages_and_clamp_each_rate(
    base_rates, column, key, value, expected,
):
    table = build_rate_table(["BJ가", "BJ나"], 45, 7).astype({column: object})
    table.loc[0, column] = value

    result = rates_by_bj(table, base_rates)

    assert result["BJ가"] == pytest.approx({**base_rates, key: expected})
    assert result["BJ나"] == base_rates


@pytest.mark.parametrize(("column", "key"), [
    (SETTLE_COL, "정산비율"),
    (SUPPORT_COL, "협력지원율"),
])
@pytest.mark.parametrize("value", [
    "",
    " \t ",
    None,
    float("nan"),
    pd.NA,
    "숫자아님",
    "50%",
    "NaN",
])
def test_invalid_edits_fall_back_to_the_matching_base_rate(
    base_rates, column, key, value,
):
    table = build_rate_table(["BJ가", "BJ나"], 60, 20).astype({column: object})
    table.loc[0, column] = value
    expected = {**base_rates, "정산비율": 0.6, "협력지원율": 0.2}

    result = rates_by_bj(table, base_rates)

    assert result["BJ가"] == {**expected, key: base_rates[key]}
    assert result["BJ나"] == expected


def test_both_missing_rates_use_base_fractions_without_converting_twice(base_rates):
    table = pd.DataFrame({
        "BJ": ["BJ가"],
        SETTLE_COL: [None],
        SUPPORT_COL: [float("nan")],
    })

    assert rates_by_bj(table, base_rates) == {"BJ가": base_rates}


def test_inputs_and_each_returned_rate_dictionary_are_independent(base_rates):
    table = build_rate_table(["BJ가", "BJ나"], 60, 15)
    table_before = table.copy(deep=True)
    base_before = base_rates.copy()

    result = rates_by_bj(table, base_rates)

    assert_frame_equal(table, table_before)
    assert base_rates == base_before
    assert result["BJ가"] is not base_rates
    assert result["BJ나"] is not base_rates
    assert result["BJ가"] is not result["BJ나"]

    result["BJ가"]["정산비율"] = 0.99
    result["BJ가"]["투네이션수수료"] = 0.88
    assert result["BJ나"] == {
        **base_before, "정산비율": 0.60, "협력지원율": 0.15,
    }
    assert base_rates == base_before
    assert_frame_equal(table, table_before)

    base_rates["협력지원율"] = 0.77
    base_rates["펀네이션수수료"] = 0.66
    for rates in result.values():
        assert rates["협력지원율"] == 0.15
        assert rates["펀네이션수수료"] == base_before["펀네이션수수료"]
