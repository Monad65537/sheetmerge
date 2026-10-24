from dataclasses import replace

import pytest
from conftest import HEADERS, VALID

from sheetmerge import consolidate, load_batch, reconcile


def test_composite_key_keeps_different_branches_and_transactions(
    tmp_path, write_csv, config_factory
):
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            VALID,
            ["002", *VALID[1:]],
            [VALID[0], "000043", *VALID[2:]],
        ],
    )
    result = consolidate(load_batch(tmp_path, config_factory()))
    assert len(result.accepted) == 3
    assert result.rejected == result.redundant == ()


def test_exact_duplicates_use_normalized_values_and_earliest_provenance(
    tmp_path, write_csv, profile_factory, config_factory
):
    write_csv(
        tmp_path / "z.csv",
        rows=[[*VALID[:4], "12.3400", "USD", "different note"]],
        headers=[*HEADERS, "note"],
    )
    write_csv(
        tmp_path / "a.csv", rows=[[*VALID, "note"], [*VALID, "other"]], headers=[*HEADERS, "note"]
    )
    config = config_factory(profile_factory(ignored_headers=("note",)))
    batch = load_batch(tmp_path, config)
    result = consolidate(replace(batch, rows=tuple(reversed(batch.rows))))
    assert len(result.accepted) == 1
    winner = result.accepted[0].row.source
    assert (winner.source_file, winner.source_row) == ("a.csv", 2)
    assert len(result.redundant) == 2
    assert all(item.duplicate_of == winner for item in result.redundant)
    assert all(item.issues[0].code == "EXACT_DUPLICATE" for item in result.redundant)
    assert reconcile(result, config).counts == (3, 1, 0, 2)


@pytest.mark.parametrize(
    "field,value",
    [
        ("amount", "12.35"),
        ("customer_id", "another-customer"),
        ("transaction_date", "2024-03-01"),
        ("currency", "PLN"),
    ],
)
def test_every_conflicting_occurrence_is_quarantined(
    tmp_path, write_csv, config_factory, field, value
):
    conflicting = VALID.copy()
    conflicting[HEADERS.index(field)] = value
    write_csv(tmp_path / "sales.csv", rows=[VALID, VALID, conflicting])
    config = config_factory()
    result = consolidate(load_batch(tmp_path, config))
    assert result.accepted == result.redundant == ()
    assert len(result.rejected) == 3
    assert all(item.issues[0].code == "CONFLICTING_TRANSACTION" for item in result.rejected)
    assert all(field in item.issues[0].message for item in result.rejected)
    assert reconcile(result, config).counts == (3, 0, 3, 0)


def test_invalid_candidate_blocks_entire_key_and_preserves_original_reasons(
    tmp_path, write_csv, config_factory
):
    write_csv(tmp_path / "sales.csv", rows=[VALID, [*VALID[:2], "bad", *VALID[3:]], VALID])
    config = config_factory()
    result = consolidate(load_batch(tmp_path, config))
    assert len(result.rejected) == 3
    assert result.accepted == result.redundant == ()
    assert all(
        "KEY_HAS_INVALID_CANDIDATE" in {issue.code for issue in item.issues}
        for item in result.rejected
    )
    assert "INVALID_DATE" in {issue.code for issue in result.rejected[1].issues}
    assert all(item.row.amount_minor == 1234 for item in result.rejected)


def test_conflict_and_invalid_candidate_reasons_are_both_retained(
    tmp_path, write_csv, config_factory
):
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            VALID,
            [*VALID[:4], "99.00", "USD"],
            [*VALID[:2], "bad", *VALID[3:]],
        ],
    )
    result = consolidate(load_batch(tmp_path, config_factory()))
    for item in result.rejected:
        assert {"CONFLICTING_TRANSACTION", "KEY_HAS_INVALID_CANDIDATE"} <= {
            issue.code for issue in item.issues
        }


def test_blank_and_missing_key_are_explicit_rejections(tmp_path, write_csv, config_factory):
    write_csv(tmp_path / "sales.csv", rows=[[""] * 6, ["", *VALID[1:]], VALID])
    config = config_factory()
    result = consolidate(load_batch(tmp_path, config))
    assert len(result.accepted) == 1
    assert {issue.code for item in result.rejected for issue in item.issues} == {
        "BLANK_ROW",
        "MISSING_REQUIRED_VALUE",
    }
    assert reconcile(result, config).counts == (3, 1, 2, 0)


def test_normalization_is_the_only_authorized_change_to_keys(tmp_path, write_csv, config_factory):
    write_csv(
        tmp_path / "sales.csv",
        rows=[
            VALID,
            ["1", *VALID[1:]],
            [" 001 ", *VALID[1:]],
            ["001", "000042A", *VALID[2:]],
        ],
    )
    result = consolidate(load_batch(tmp_path, config_factory()))
    assert len(result.accepted) == 4
    assert not result.redundant


def test_duplicates_and_conflicts_span_csv_and_xlsx_sheets(
    tmp_path, write_csv, write_xlsx, config_factory, profile_factory
):
    write_csv(tmp_path / "a.csv")
    write_xlsx(tmp_path / "z.xlsx", sheets={"Z": [VALID], "A": [VALID]})
    config = config_factory(profile_factory(), profile_factory("xlsx", sheets=("Z", "A")))
    result = consolidate(load_batch(tmp_path, config))
    assert len(result.redundant) == 2
    assert result.accepted[0].row.source.source_file == "a.csv"
    write_xlsx(tmp_path / "z.xlsx", sheets={"Z": [VALID], "A": [[*VALID[:4], "12.35", "USD"]]})
    result = consolidate(load_batch(tmp_path, config))
    assert len(result.rejected) == 3
    assert result.accepted == result.redundant == ()
