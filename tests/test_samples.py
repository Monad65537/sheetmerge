"""The shipped synthetic files are an executable example of the documented policies."""

import hashlib
import json
from collections import Counter
from pathlib import Path

import pytest

from sheetmerge import (
    IngestionError,
    consolidate,
    discover_inputs,
    load_batch,
    load_config,
    reconcile,
)
from sheetmerge.cli import main

SAMPLES = Path(__file__).parents[1] / "samples"


def test_sample_inventory_and_explicit_profiles():
    files = discover_inputs(SAMPLES / "input")
    assert len(files) == 12
    assert Counter(path.suffix for path in files) == {".csv": 6, ".xlsx": 6}
    config = load_config(SAMPLES / "sources.toml")
    for path in files:
        relative = path.relative_to(SAMPLES / "input").as_posix()
        assert config.profile_for(relative).format == path.suffix[1:]


def test_processable_samples_match_documented_outcomes():
    expected = json.loads((SAMPLES / "expected.json").read_text())
    config = load_config(SAMPLES / "sources.toml")
    batch = load_batch(SAMPLES / "input" / "processable", config)
    assert len(batch.files) == expected["file_counts"]["processable"]
    result = consolidate(batch)
    report = reconcile(result, config).to_dict()
    assert report["counts"] == expected["counts"]
    assert report["unquantified_count"] == expected["unquantified_count"]
    assert report["unknown_currency_count"] == expected["unknown_currency_count"]
    assert report["monetary_complete"] is expected["monetary_complete"]
    assert {item["currency"]: item["amounts"] for item in report["currencies"]} == (
        expected["currencies"]
    )
    assert (
        Counter(issue.code for item in result.rows for issue in item.issues)
        == (expected["reason_counts"])
    )
    for row in batch.rows:
        assert row.values["branch_id"] in ("001", "002", "003")
        for field in ("transaction_id", "customer_id"):
            assert isinstance(row.values[field], str)
            assert row.values[field].startswith("000")
        assert row.source.source_row in (2, 3, 4)
        if row.source.source_sheet:
            assert row.source.source_sheet == "Sales"
    # The repeated transaction ID across three branches is not one shared key.
    assert {
        item.row.values["branch_id"]
        for item in result.accepted
        if item.row.values["transaction_id"] == "000001"
    } == {"001", "002", "003"}


def test_missing_column_fixture_is_a_deliberate_structural_failure():
    config = load_config(SAMPLES / "sources.toml")
    with pytest.raises(IngestionError) as error:
        load_batch(SAMPLES / "input", config)
    assert "branch_c_04_missing_columns.xlsx" in str(error.value)
    assert "account_id" in str(error.value)
    assert "currency_code" in str(error.value)


def test_sample_cli_runs_preserve_every_source(tmp_path, capsys):
    files = discover_inputs(SAMPLES / "input")
    before = {
        path: (hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns)
        for path in files
    }
    common = ["--config", str(SAMPLES / "sources.toml")]
    assert (
        main(
            [
                *common,
                "--input",
                str(SAMPLES / "input"),
                "--output",
                str(tmp_path / "structural-failure"),
            ]
        )
        == 2
    )
    assert not (tmp_path / "structural-failure").exists()
    assert "branch_c_04_missing_columns.xlsx" in capsys.readouterr().err
    assert (
        main(
            [
                *common,
                "--input",
                str(SAMPLES / "input" / "processable"),
                "--output",
                str(tmp_path / "review"),
            ]
        )
        == 1
    )
    summary = json.loads(capsys.readouterr().out)
    assert summary["counts"] == {"input": 33, "accepted": 16, "rejected": 14, "redundant": 3}
    assert summary["status"] == "review_required"
    assert (tmp_path / "review" / "master.xlsx").is_file()
    after = {
        path: (hashlib.sha256(path.read_bytes()).hexdigest(), path.stat().st_mtime_ns)
        for path in files
    }
    assert after == before
