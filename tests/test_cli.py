import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from conftest import VALID

from sheetmerge import ReconciliationError, cli


def arguments(tmp_path, write_config):
    config = write_config(tmp_path / "sources.toml")
    return [
        "--config",
        str(config),
        "--input",
        str(tmp_path / "inputs"),
        "--output",
        str(tmp_path / "out"),
    ]


def test_cli_clean_batch_and_exact_duplicates_exit_zero(tmp_path, write_config, write_csv, capsys):
    write_csv(tmp_path / "inputs" / "sales.csv", rows=[VALID, VALID])
    assert cli.main(arguments(tmp_path, write_config)) == 0
    captured = capsys.readouterr()
    summary = json.loads(captured.out)
    assert summary["status"] == "completed"
    assert summary["counts"] == {"input": 2, "accepted": 1, "rejected": 0, "redundant": 1}
    assert summary["monetary_complete"] is True
    assert captured.err == ""
    assert (tmp_path / "out" / "master.xlsx").exists()


@pytest.mark.parametrize(
    "rows,complete",
    [
        ([VALID, [*VALID[:4], "99.00", "USD"]], True),
        ([[*VALID[:4], "invalid", "USD"]], False),
        ([[""] * 6], False),
    ],
)
def test_cli_rejections_and_incomplete_money_require_review(
    tmp_path, write_config, write_csv, capsys, rows, complete
):
    write_csv(tmp_path / "inputs" / "sales.csv", rows=rows)
    assert cli.main(arguments(tmp_path, write_config)) == 1
    summary = json.loads(capsys.readouterr().out)
    assert summary["status"] == "review_required"
    assert summary["monetary_complete"] is complete
    assert summary["counts"]["accepted"] == 0
    assert summary["counts"]["rejected"] == len(rows)
    assert (tmp_path / "out" / "exclusions.csv").exists()


def test_cli_bad_configuration_fails_clearly_without_output(
    tmp_path, write_config, write_csv, capsys
):
    write_csv(tmp_path / "inputs" / "sales.csv")
    args = arguments(tmp_path, write_config)
    (tmp_path / "sources.toml").write_text("not valid TOML")
    assert cli.main(args) == 2
    captured = capsys.readouterr()
    assert "sheetmerge: error:" in captured.err
    assert captured.out == ""
    assert not (tmp_path / "out").exists()


def test_cli_reconciliation_failure_is_fatal(
    tmp_path, write_config, write_csv, capsys, monkeypatch
):
    write_csv(tmp_path / "inputs" / "sales.csv")

    def fail(*args):
        raise ReconciliationError("count mismatch")

    monkeypatch.setattr(cli, "export_batch", fail)
    assert cli.main(arguments(tmp_path, write_config)) == 2
    captured = capsys.readouterr()
    assert "count mismatch" in captured.err
    assert captured.out == ""
    assert not (tmp_path / "out").exists()


def test_cli_structural_failure_is_fatal(tmp_path, write_config, write_csv, capsys):
    write_csv(tmp_path / "inputs" / "sales.csv", rows=[["too", "few", "columns"]])
    assert cli.main(arguments(tmp_path, write_config)) == 2
    assert "sales.csv" in capsys.readouterr().err
    assert not (tmp_path / "out").exists()


def test_python_module_entry_point_handles_mixed_inputs(
    tmp_path, write_config, write_csv, write_xlsx
):
    inputs = tmp_path / "inputs"
    write_csv(inputs / "sales.csv")
    write_xlsx(inputs / "sales.xlsx", rows=[["002", *VALID[1:]]])
    env = os.environ.copy()
    env["PYTHONPATH"] = str(Path(__file__).parents[1] / "src")
    process = subprocess.run(
        [sys.executable, "-m", "sheetmerge", *arguments(tmp_path, write_config)],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )
    assert process.returncode == 0, process.stderr
    assert json.loads(process.stdout)["counts"]["accepted"] == 2
    assert process.stderr == ""


def test_cli_argument_validation():
    with pytest.raises(SystemExit) as error:
        cli.main([])
    assert error.value.code == 2
