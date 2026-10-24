from dataclasses import replace
from decimal import localcontext

import pytest
from conftest import VALID

from sheetmerge import (
    ClassifiedRow,
    ReconciliationError,
    consolidate,
    load_batch,
    reconcile,
)
from sheetmerge.models import Issue


def test_counts_money_and_unknown_values_are_accounted_for(tmp_path, write_csv, config_factory):
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            ["001", "A", VALID[2], "C", "0.10", "USD"],
            ["001", "A", VALID[2], "C", "0.10", "USD"],
            ["001", "P", VALID[2], "C", "10.20", "PLN"],
            ["001", "B", VALID[2], "C", "2.00", "USD"],
            ["001", "B", VALID[2], "C", "3.00", "USD"],
            ["001", "C", "invalid", "C", "1.11", "USD"],
            ["001", "D", VALID[2], "C", "invalid", "USD"],
            ["001", "E", VALID[2], "C", "5.00", "EUR"],
            [""] * 6,
        ],
    )
    config = config_factory()
    report = reconcile(consolidate(load_batch(tmp_path, config)), config)
    assert report.counts == (9, 2, 6, 1)
    currencies = {item.currency: item for item in report.currencies}
    assert currencies["USD"].counts == (6, 1, 4, 1)
    assert currencies["USD"].totals == (631, 10, 611, 10)
    assert currencies["USD"].unquantified == (1, 0, 1, 0)
    assert currencies["PLN"].totals == (1020, 1020, 0, 0)
    assert report.unknown_currency_count == 2
    assert report.unquantified_count == 3
    assert not report.complete
    data = report.to_dict()
    assert data["count_balanced"] is True
    assert data["known_amounts_balanced"] is True
    assert data["monetary_complete"] is False
    assert currencies["USD"].to_dict()["amounts"]["input"] == "6.31"


def test_large_signed_totals_use_integer_arithmetic(tmp_path, write_csv, config_factory):
    large = "999999999999999999999999999999999999.99"
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            ["001", "A", VALID[2], "C", large, "USD"],
            ["001", "B", VALID[2], "C", large, "USD"],
            ["001", "C", VALID[2], "C", "-0.01", "USD"],
            ["001", "D", VALID[2], "C", "0.00", "USD"],
        ],
    )
    config = config_factory()
    with localcontext() as context:
        context.prec = 2
        report = reconcile(consolidate(load_batch(tmp_path, config)), config)
        usd = next(item for item in report.currencies if item.currency == "USD")
        assert usd.totals[0] == 2 * (10**38 - 1) - 1
        assert usd.to_dict()["amounts"]["input"] == "1999999999999999999999999999999999999.97"
    assert report.complete


@pytest.fixture
def valid_result(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv", rows=[VALID, VALID])
    config = config_factory()
    return consolidate(load_batch(tmp_path, config)), config


@pytest.mark.parametrize("mutation", ["lost", "duplicated", "unknown_status", "changed"])
def test_reconciliation_detects_broken_record_partition(valid_result, mutation):
    result, config = valid_result
    rows = list(result.rows)
    if mutation == "lost":
        rows.pop()
    elif mutation == "duplicated":
        rows.append(rows[0])
    elif mutation == "unknown_status":
        rows[0] = replace(rows[0], status="unknown")
    else:
        row = replace(rows[0].row, values={**rows[0].row.values, "customer_id": "changed"})
        rows[0] = replace(rows[0], row=row)
    with pytest.raises(ReconciliationError):
        reconcile(replace(result, rows=tuple(rows)), config)


def test_incorrect_minor_unit_metadata_fails_even_when_both_ledgers_match(valid_result):
    result, config = valid_result
    bad_row = replace(result.rows[0].row, amount_minor=999)
    batch = replace(result.batch, rows=(bad_row, result.batch.rows[1]))
    result = replace(
        result, batch=batch, rows=(replace(result.rows[0], row=bad_row), result.rows[1])
    )
    with pytest.raises(ReconciliationError, match="minor-unit"):
        reconcile(result, config)


def test_conflict_cannot_be_disguised_as_a_redundant_record(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv", rows=[VALID, [*VALID[:4], "99.00", "USD"]])
    config = config_factory()
    result = consolidate(load_batch(tmp_path, config))
    a, b = result.rows
    forged = replace(
        result,
        rows=(
            ClassifiedRow(a.row, "accepted"),
            ClassifiedRow(
                b.row, "redundant", (Issue("EXACT_DUPLICATE", "row", "forged"),), a.row.source
            ),
        ),
    )
    with pytest.raises(ReconciliationError, match="not identical"):
        reconcile(forged, config)


def test_accepting_one_member_of_quarantine_is_detected(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv", rows=[VALID, [*VALID[:2], "invalid", *VALID[3:]]])
    config = config_factory()
    result = consolidate(load_batch(tmp_path, config))
    rows = (ClassifiedRow(result.rows[0].row, "accepted"), result.rows[1])
    with pytest.raises(ReconciliationError, match="quarantined candidate"):
        reconcile(replace(result, rows=rows), config)


def test_missing_exclusion_reason_is_detected(valid_result):
    result, config = valid_result
    with pytest.raises(ReconciliationError, match="no reason"):
        reconcile(
            replace(result, rows=(result.rows[0], replace(result.rows[1], issues=()))), config
        )


def test_accepting_duplicate_keys_is_detected(valid_result):
    result, config = valid_result
    rows = tuple(ClassifiedRow(item.row, "accepted") for item in result.rows)
    with pytest.raises(ReconciliationError, match="duplicate business key"):
        reconcile(replace(result, rows=rows), config)


def test_discarding_identical_valid_group_is_detected(valid_result):
    result, config = valid_result
    rows = tuple(
        ClassifiedRow(item.row, "rejected", (Issue("WRONG", "row", "forged"),))
        for item in result.rows
    )
    with pytest.raises(ReconciliationError, match="nonconflicting"):
        reconcile(replace(result, rows=rows), config)


def test_nondeterministic_duplicate_representative_is_detected(valid_result):
    result, config = valid_result
    a, b = result.rows
    rows = (
        ClassifiedRow(a.row, "redundant", b.issues, b.row.source),
        ClassifiedRow(b.row, "accepted"),
    )
    with pytest.raises(ReconciliationError, match="not deterministic"):
        reconcile(replace(result, rows=rows), config)


def test_empty_dataset_balances(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "empty.csv", rows=[])
    config = config_factory()
    report = reconcile(consolidate(load_batch(tmp_path, config)), config)
    assert report.counts == (0, 0, 0, 0)
    assert report.complete


@pytest.mark.parametrize("mutation", ["duplicate_inventory", "missing_snapshot"])
def test_provenance_requires_unique_input_snapshots(valid_result, mutation):
    result, config = valid_result
    files = result.batch.files * 2 if mutation == "duplicate_inventory" else ()
    batch = replace(result.batch, files=files)
    with pytest.raises(ReconciliationError, match="inventory|snapshot"):
        reconcile(replace(result, batch=batch), config)
