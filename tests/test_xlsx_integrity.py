"""Regression cases for ambiguous worksheet coordinates and lossy date decoding."""

import copy
import csv
import io
import json
from datetime import date
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from conftest import VALID

from sheetmerge import IngestionError, consolidate, export_batch, ingest, load_batch
from sheetmerge.cli import main

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
PART = "xl/worksheets/sheet1.xml"


def edit_xml(path, edit, member=PART):
    """Prepare a synthetic malformed fixture, before exercising read-only ingestion."""
    output = io.BytesIO()
    with ZipFile(io.BytesIO(path.read_bytes())) as source, ZipFile(output, "w") as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == member:
                root = ET.fromstring(data)
                edit(root)
                data = ET.tostring(root, encoding="utf-8").replace(b"\r", b"&#13;")
            target.writestr(info, data)
    path.write_bytes(output.getvalue())


def set_iso_date(path, text, column="C"):
    def edit(root):
        cell = root.find(f"{NS}sheetData/{NS}row[@r='2']/{NS}c[@r='{column}2']")
        cell.clear()
        cell.set("r", f"{column}2")
        cell.set("t", "d")
        ET.SubElement(cell, f"{NS}v").text = text

    edit_xml(path, edit)


@pytest.fixture
def workbook_input(tmp_path, write_xlsx, profile_factory, config_factory):
    path = write_xlsx(
        tmp_path / "sales.xlsx",
        rows=[
            [*VALID[:4], 1, "USD"],
            ["001", "T2", VALID[2], "C2", 2, "USD"],
        ],
    )
    config = config_factory(profile_factory("xlsx", allow_excel_dates=True))
    return path, config


@pytest.mark.parametrize(
    "fault",
    [
        "duplicate_row",
        "out_of_order_rows",
        "duplicate_cell",
        "cell_row_mismatch",
        "out_of_order_cells",
        "bad_cell_address",
        "column_overflow",
        "row_zero",
        "row_overflow",
        "nonnumeric_row",
        "duplicate_value",
        "missing_sheet_data",
        "duplicate_sheet_data",
        "orphan_cell",
        "unsupported_cell_type",
    ],
)
def test_malformed_coordinates_and_structure_fail_without_changing_source(workbook_input, fault):
    path, config = workbook_input

    def edit(root):
        table = root.find(f"{NS}sheetData")
        row = table[1]
        if fault == "duplicate_row":
            table[2].set("r", "2")
            for cell in table[2]:
                cell.set("r", cell.get("r").replace("3", "2"))
        elif fault == "out_of_order_rows":
            table[:] = [table[0], table[2], table[1]]
        elif fault == "duplicate_cell":
            duplicate = copy.deepcopy(row[4])
            duplicate.find(f"{NS}v").text = "999.99"
            row.append(duplicate)
        elif fault == "cell_row_mismatch":
            row[0].set("r", "A3")
        elif fault == "out_of_order_cells":
            row[0], row[1] = row[1], row[0]
        elif fault == "bad_cell_address":
            row[0].set("r", "$A$2")
        elif fault == "column_overflow":
            row[-1].set("r", "XFE2")
        elif fault in ("row_zero", "row_overflow", "nonnumeric_row"):
            row.set("r", {"row_zero": "0", "row_overflow": "1048577", "nonnumeric_row": "x"}[fault])
        elif fault == "duplicate_value":
            row[4].append(copy.deepcopy(row[4].find(f"{NS}v")))
        elif fault == "missing_sheet_data":
            root.remove(table)
        elif fault == "duplicate_sheet_data":
            root.append(copy.deepcopy(table))
        elif fault == "orphan_cell":
            root.append(copy.deepcopy(row[0]))
        elif fault == "unsupported_cell_type":
            row[0].set("t", "unknown")

    edit_xml(path, edit)
    before = path.read_bytes(), path.stat().st_mtime_ns
    with pytest.raises(IngestionError) as error:
        load_batch(path.parent, config)
    assert "sales.xlsx" in str(error.value)
    assert "sheet 'Sales'" in str(error.value)
    assert (path.read_bytes(), path.stat().st_mtime_ns) == before


def test_coordinate_preflight_runs_before_value_decoding(workbook_input, monkeypatch):
    path, config = workbook_input
    edit_xml(path, lambda root: root.find(f"{NS}sheetData")[2].set("r", "2"))

    def unexpected_decode(*args, **kwargs):
        pytest.fail("openpyxl must not decode an invalid worksheet")

    monkeypatch.setattr(ingest, "load_workbook", unexpected_decode)
    with pytest.raises(IngestionError, match="duplicate or out-of-order row"):
        load_batch(path.parent, config)


def test_duplicate_zip_names_fail(workbook_input):
    path, config = workbook_input
    with ZipFile(path) as archive:
        duplicate = archive.read(PART)
    with pytest.warns(UserWarning, match="Duplicate name"), ZipFile(path, "a") as archive:
        archive.writestr(PART, duplicate)
    with pytest.raises(IngestionError, match="Duplicate ZIP"):
        load_batch(path.parent, config)


def test_duplicate_relationship_ids_fail(workbook_input):
    path, config = workbook_input
    edit_xml(path, lambda root: root.append(copy.deepcopy(root[0])), "xl/_rels/workbook.xml.rels")
    with pytest.raises(IngestionError, match="Duplicate workbook relationship"):
        load_batch(path.parent, config)


def test_sparse_rows_keep_their_physical_numbers_and_original_amounts(workbook_input):
    path, config = workbook_input

    def edit(root):
        row = root.find(f"{NS}sheetData")[2]
        row.set("r", "5")
        for cell in row:
            cell.set("r", cell.get("r").replace("3", "5"))

    edit_xml(path, edit)
    batch = load_batch(path.parent, config)
    assert [row.source.source_row for row in batch.rows] == [2, 3, 4, 5]
    assert [row.amount_minor for row in batch.accepted] == [100, 200]
    assert [row.values["transaction_id"] for row in batch.accepted] == [VALID[1], "T2"]
    assert len(batch.blank) == 2


def test_unambiguous_implicit_coordinates_remain_supported(workbook_input):
    path, config = workbook_input

    def edit(root):
        for row in root.find(f"{NS}sheetData"):
            del row.attrib["r"]
            for cell in row:
                del cell.attrib["r"]

    edit_xml(path, edit)
    batch = load_batch(path.parent, config)
    assert [row.source.source_row for row in batch.accepted] == [2, 3]
    assert [row.amount_minor for row in batch.accepted] == [100, 200]


def test_coordinate_validation_does_not_normalize_identifier_content(workbook_input):
    path, config = workbook_input

    def edit(root):
        text = root.find(f"{NS}sheetData/{NS}row[@r='2']/{NS}c[@r='D2']/{NS}is/{NS}t")
        text.text = "000\r\n007"

    edit_xml(path, edit)
    assert load_batch(path.parent, config).accepted[0].values["customer_id"] == "000\r\n007"


@pytest.mark.parametrize(
    "text",
    [
        "2026-01-01garbage",
        "2026-01-01T00:00:00.000001",
        "2026-01-01T00:00:00.0000001",
        "2026-01-01T12:00:00",
        "2026-01-01T00:00:00Z",
        "2026-01-01T00:00:00+00:00",
        "2026-01-01 ",
        " 2026-01-01",
        "2026-1-01",
        "2026-02-30",
        "0000-01-01",
        "00:00:00",
        "2026-01-01T00:00",
        "2026-01-01T00:00:00.",
    ],
)
def test_native_iso_dates_are_strictly_rejected_and_originals_retained(workbook_input, text):
    path, config = workbook_input
    set_iso_date(path, text)
    batch = load_batch(path.parent, config)
    row = batch.rejected[0]
    assert [issue.code for issue in row.issues] == ["INVALID_DATE"]
    assert row.values["transaction_date"] is None
    assert row.raw_values["transaction_date"] == text
    assert row.raw_cells["transaction_date"].original_text == text
    assert len(batch.accepted) == 1  # The other valid transaction is unaffected.


@pytest.mark.parametrize(
    "text",
    [
        "2026-01-01",
        "2026-01-01T00:00:00",
        "2026-01-01T00:00:00.0",
        "2026-01-01T00:00:00.000000000",
        "0001-01-01",
        "9999-12-31",
    ],
)
def test_supported_native_iso_dates_remain_valid(workbook_input, text):
    path, config = workbook_input
    set_iso_date(path, text)
    row = load_batch(path.parent, config).accepted[0]
    assert row.values["transaction_date"] == date.fromisoformat(text[:10])
    assert row.raw_values["transaction_date"] == text


def test_native_iso_date_opt_in_remains_required(workbook_input, profile_factory, config_factory):
    path, _ = workbook_input
    set_iso_date(path, "2026-01-01")
    row = load_batch(path.parent, config_factory(profile_factory("xlsx"))).rejected[0]
    assert row.issues[0].code == "INVALID_DATE"
    assert row.raw_values["transaction_date"] == "2026-01-01"


@pytest.mark.parametrize(
    "column,text,reason",
    [
        ("A", "001", "ID_NOT_TEXT"),
        ("D", "000007", "ID_NOT_TEXT"),
        ("E", "12.34", "INVALID_AMOUNT"),
        ("F", "USD", "UNSUPPORTED_CURRENCY"),
    ],
)
def test_native_date_cells_cannot_be_reinterpreted_as_other_field_types(
    workbook_input, column, text, reason
):
    path, config = workbook_input
    set_iso_date(path, text, column)
    row = load_batch(path.parent, config).rejected[0]
    assert row.issues[0].code == reason


def test_date_originals_survive_exclusion_and_audit_exports(workbook_input, tmp_path):
    path, config = workbook_input
    original = "2026-01-01T00:00:00.000001"
    set_iso_date(path, original)
    before = path.read_bytes(), path.stat().st_mtime_ns
    batch = load_batch(path.parent, config)
    export_batch(consolidate(batch), config, tmp_path.parent / f"{tmp_path.name}-output")
    output = tmp_path.parent / f"{tmp_path.name}-output"
    with (output / "exclusions.csv").open(encoding="utf-8", newline="") as stream:
        excluded = list(csv.DictReader(stream))
    assert json.loads(excluded[0]["raw_values"])["transaction_date"] == original
    audit = [json.loads(line) for line in (output / "audit.jsonl").read_text().splitlines()]
    rejected = next(item for item in audit if item["status"] == "rejected")
    assert rejected["raw_cells"]["transaction_date"]["original_text"] == original
    assert (path.read_bytes(), path.stat().st_mtime_ns) == before


def test_shared_string_headers_and_identifiers_survive_preflight(workbook_input):
    path, config = workbook_input

    def convert_cells(root):
        for reference, index in (("A1", "0"), ("B2", "1")):
            cell = root.find(f"{NS}sheetData/{NS}row/{NS}c[@r='{reference}']")
            cell.clear()
            cell.set("r", reference)
            cell.set("t", "s")
            ET.SubElement(cell, f"{NS}v").text = index

    edit_xml(path, convert_cells)
    edit_xml(
        path,
        lambda root: ET.SubElement(
            root,
            "{http://schemas.openxmlformats.org/package/2006/content-types}Override",
            PartName="/xl/sharedStrings.xml",
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        ),
        "[Content_Types].xml",
    )
    edit_xml(
        path,
        lambda root: ET.SubElement(
            root,
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
            Id="rIdSharedStrings",
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            Target="sharedStrings.xml",
        ),
        "xl/_rels/workbook.xml.rels",
    )
    shared = ET.Element(f"{NS}sst", count="2", uniqueCount="2")
    for value in ("branch_id", VALID[1]):
        ET.SubElement(ET.SubElement(shared, f"{NS}si"), f"{NS}t").text = value
    with ZipFile(path, "a") as archive:
        archive.writestr("xl/sharedStrings.xml", ET.tostring(shared, encoding="utf-8"))
    batch = load_batch(path.parent, config)
    assert len(batch.accepted) == 2
    assert batch.accepted[0].values["transaction_id"] == VALID[1]
    assert [row.amount_minor for row in batch.accepted] == [100, 200]


def test_numeric_date_rejection_retains_original_serial(
    tmp_path, write_xlsx, profile_factory, config_factory
):
    path = write_xlsx(
        tmp_path / "sales.xlsx",
        rows=[
            [*VALID[:2], date(2026, 1, 1), *VALID[3:]],
        ],
    )
    serial = "61.000000000001"
    edit_xml(
        path,
        lambda root: setattr(
            root.find(f"{NS}sheetData/{NS}row[@r='2']/{NS}c[@r='C2']/{NS}v"), "text", serial
        ),
    )
    config = config_factory(profile_factory("xlsx", allow_excel_dates=True))
    row = load_batch(tmp_path, config).rejected[0]
    assert row.issues[0].code == "INVALID_DATE"
    assert row.raw_values["transaction_date"] == serial
    assert row.raw_cells["transaction_date"].numeric_text == serial
    assert row.raw_cells["transaction_date"].original_text == serial


def test_coordinate_failure_is_fatal_to_cli_and_does_not_publish(
    workbook_input, write_config, capsys
):
    path, _ = workbook_input
    edit_xml(path, lambda root: root.find(f"{NS}sheetData")[2].set("r", "2"))
    config_file = write_config(path.parent / "sources.toml")
    output = path.parent.parent / f"{path.parent.name}-output"
    assert (
        main(["--input", str(path.parent), "--config", str(config_file), "--output", str(output)])
        == 2
    )
    captured = capsys.readouterr()
    assert "duplicate or out-of-order row" in captured.err
    assert captured.out == ""
    assert not output.exists()


def test_unselected_worksheets_remain_excluded(
    tmp_path, write_xlsx, profile_factory, config_factory
):
    path = write_xlsx(
        tmp_path / "sales.xlsx",
        sheets={
            "Sales": [VALID],
            "Ignored": [VALID, VALID],
        },
    )
    edit_xml(
        path, lambda root: root.find(f"{NS}sheetData")[2].set("r", "2"), "xl/worksheets/sheet2.xml"
    )
    batch = load_batch(tmp_path, config_factory(profile_factory("xlsx")))
    assert len(batch.accepted) == 1
    assert batch.accepted[0].source.source_sheet == "Sales"
