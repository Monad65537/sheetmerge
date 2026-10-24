from decimal import Decimal, localcontext

import pytest
from conftest import HEADERS, VALID

from sheetmerge import load_batch


@pytest.mark.parametrize(
    "field,value,code",
    [
        ("branch_id", "", "MISSING_REQUIRED_VALUE"),
        ("transaction_id", "   ", "MISSING_REQUIRED_VALUE"),
        ("customer_id", "", "MISSING_REQUIRED_VALUE"),
        ("transaction_date", "", "MISSING_REQUIRED_VALUE"),
        ("amount", "", "MISSING_REQUIRED_VALUE"),
        ("currency", "", "MISSING_REQUIRED_VALUE"),
        ("transaction_date", "2023-02-29", "INVALID_DATE"),
        ("transaction_date", "2024-2-29", "INVALID_DATE"),
        ("transaction_date", "29/02/2024", "INVALID_DATE"),
        ("transaction_date", "2024-02-29T12:00:00", "INVALID_DATE"),
        ("amount", "NaN", "INVALID_AMOUNT"),
        ("amount", "Infinity", "INVALID_AMOUNT"),
        ("amount", "1e2", "INVALID_AMOUNT"),
        ("amount", "1,234.50", "INVALID_AMOUNT"),
        ("amount", "0.001", "EXCESS_FRACTIONAL_PRECISION"),
        ("amount", "12.34 ", "INVALID_AMOUNT"),
        ("amount", "1000000000000000000000000000000000000", "AMOUNT_TOO_LARGE"),
        ("currency", "usd", "UNSUPPORTED_CURRENCY"),
        ("currency", "EUR", "UNSUPPORTED_CURRENCY"),
    ],
)
def test_field_rejection_reasons(tmp_path, write_csv, config_factory, field, value, code):
    record = VALID.copy()
    record[HEADERS.index(field)] = value
    write_csv(tmp_path / "sales.csv", rows=[record])
    result = load_batch(tmp_path, config_factory())
    assert not result.accepted
    assert len(result.rows) == len(result.rejected) == 1
    row = result.rejected[0]
    assert [(issue.field, issue.code) for issue in row.issues] == [(field, code)]
    assert row.source.source_row == 2
    assert row.raw_values[field] == value


def test_all_independent_failures_are_reported_and_known_money_retained(
    tmp_path, write_csv, config_factory
):
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            ["", "", "impossible", "", "0.10", "USD"],
        ],
    )
    row = load_batch(tmp_path, config_factory()).rejected[0]
    assert len(row.issues) == 4
    assert row.amount_minor == 10


@pytest.mark.parametrize(
    "amount,currency,expected",
    [
        ("0.10", "USD", 10),
        ("-0.01", "USD", -1),
        ("+12.34", "USD", 1234),
        ("0", "USD", 0),
        ("-0.00", "USD", 0),
        ("1.2300", "USD", 123),
        ("12", "JPY", 12),
        ("1.234", "KWD", 1234),
        ("123456789012345678901234567890.12", "USD", 12345678901234567890123456789012),
    ],
)
def test_exact_money_independent_of_decimal_context(
    tmp_path, write_csv, config_factory, amount, currency, expected
):
    write_csv(tmp_path / "sales.csv", rows=[[*VALID[:4], amount, currency]])
    with localcontext() as context:
        context.prec = 2
        result = load_batch(tmp_path, config_factory())
    assert result.accepted[0].amount_minor == expected
    assert result.accepted[0].values["amount"] == Decimal(amount)
    assert result.to_frame().loc[0, "amount_minor"] == expected


def test_integer_money_can_be_summed_exactly(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv", rows=[[*VALID[:4], "0.10", "USD"]] * 10)
    result = load_batch(tmp_path, config_factory())
    assert sum(row.amount_minor for row in result.accepted) == 100
    assert len(result.accepted) == 10


@pytest.mark.parametrize("amount", ["12,34.50", "1,23,456.00", "1,,234", "1,234."])
def test_grouping_positions_are_validated(
    tmp_path, write_csv, profile_factory, config_factory, amount
):
    write_csv(tmp_path / "sales.csv", rows=[[*VALID[:4], amount, "USD"]])
    profile = profile_factory(grouping_separator=",")
    row = load_batch(tmp_path, config_factory(profile)).rejected[0]
    assert row.issues[0].code == "INVALID_AMOUNT"


def test_ambiguous_configured_date_is_rejected(
    tmp_path, write_csv, profile_factory, config_factory
):
    write_csv(tmp_path / "sales.csv", rows=[[*VALID[:2], "01/02/2024", *VALID[3:]]])
    profile = profile_factory(date_formats=("%d/%m/%Y", "%m/%d/%Y"))
    row = load_batch(tmp_path, config_factory(profile)).rejected[0]
    assert row.issues[0].code == "INVALID_DATE"
    assert "different interpretations" in row.issues[0].message


def test_ids_preserve_whitespace_case_unicode_and_long_digits(tmp_path, write_csv, config_factory):
    identifiers = [" 00Ab ", "000000000000000000000000000001", "Łódź-007"]
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            [identifiers[0], identifiers[1], VALID[2], identifiers[2], *VALID[4:]],
        ],
    )
    row = load_batch(tmp_path, config_factory()).accepted[0]
    assert [row.values[field] for field in ("branch_id", "transaction_id", "customer_id")] == (
        identifiers
    )


def test_invalid_currency_retains_parseable_decimal_but_no_minor_units(
    tmp_path, write_csv, config_factory
):
    write_csv(tmp_path / "sales.csv", rows=[[*VALID[:5], "EUR"]])
    row = load_batch(tmp_path, config_factory()).rejected[0]
    assert row.values["amount"] == Decimal("12.34")
    assert row.amount_minor is None


def test_blank_rows_are_distinct_from_partial_rows(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv", rows=[[""] * 6, ["001", "", "", "", "", ""]])
    result = load_batch(tmp_path, config_factory())
    assert len(result.blank) == len(result.rejected) == 1
    assert len(result.rejected[0].issues) == 5


def test_audit_values_are_immutable(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv")
    row = load_batch(tmp_path, config_factory()).accepted[0]
    with pytest.raises(TypeError):
        row.values["branch_id"] = "changed"
    with pytest.raises(TypeError):
        row.raw_values["branch_id"] = "changed"
