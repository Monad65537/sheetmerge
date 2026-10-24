"""Small fixtures exercise configurable budgets without allocating large inputs."""

import io
import json
import re
from dataclasses import replace
from types import SimpleNamespace
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

import pytest
from conftest import VALID

from sheetmerge import (
    ConfigurationError,
    IngestionError,
    InputLimits,
    ReconciliationError,
    consolidate,
    export_batch,
    ingest,
    load_batch,
    load_config,
)
from sheetmerge.cli import main
from sheetmerge.config import LimitExceeded
from sheetmerge.xlsx import _checked_members, _parse_xml

PART = "xl/worksheets/sheet1.xml"


def limited(config, **options):
    return replace(config, limits=replace(config.limits, **options))


def rewrite_package(path, changes=None, extra=None, compression=ZIP_STORED):
    output = io.BytesIO()
    with ZipFile(io.BytesIO(path.read_bytes())) as source, ZipFile(output, "w") as target:
        for entry in source.infolist():
            target.writestr(entry, (changes or {}).get(entry.filename, source.read(entry.filename)))
        for name, data in (extra or {}).items():
            target.writestr(name, data, compress_type=compression)
    path.write_bytes(output.getvalue())


@pytest.fixture
def xlsx_input(tmp_path, write_xlsx, profile_factory, config_factory):
    path = write_xlsx(tmp_path / "sales.xlsx")
    return path, config_factory(profile_factory("xlsx"))


def test_limits_are_optional_resolved_and_configurable(tmp_path, write_config):
    path = write_config(tmp_path / "sources.toml")
    default = load_config(path)
    assert default.limits == InputLimits()
    path.write_text(path.read_text() + "\n[limits]\nmax_total_records=7\nmax_file_bytes=4096\n")
    configured = load_config(path)
    assert configured.limits.max_total_records == 7
    assert configured.limits.max_file_bytes == 4096
    assert configured.limits.max_xml_depth == default.limits.max_xml_depth
    assert configured.fingerprint != default.fingerprint


@pytest.mark.parametrize("value", [0, -1, True, 1.5, "1024"])
def test_limits_require_positive_integers(value):
    with pytest.raises(ConfigurationError, match="positive integer"):
        InputLimits(max_file_bytes=value)


@pytest.mark.parametrize(
    "options",
    [
        {"max_worksheet_rows": 1_048_577},
        {"max_worksheet_columns": 16_385},
    ],
)
def test_geometry_cannot_exceed_excel_bounds(options):
    with pytest.raises(ConfigurationError, match="physical dimensions"):
        InputLimits(**options)


def test_unknown_limit_options_are_not_ignored(tmp_path, write_config):
    path = write_config(tmp_path / "sources.toml")
    path.write_text(path.read_text() + "\n[limits]\nmax_file_byte=123\n")
    with pytest.raises(ConfigurationError, match="max_file_byte"):
        load_config(path)


@pytest.mark.parametrize("format", ["csv", "xlsx"])
def test_file_size_boundary_and_unchanged_inputs(
    tmp_path, write_csv, write_xlsx, config_factory, profile_factory, format
):
    path = (write_csv if format == "csv" else write_xlsx)(tmp_path / f"sales.{format}")
    config = config_factory(profile_factory(format))
    before = path.read_bytes(), path.stat().st_mtime_ns
    assert len(load_batch(tmp_path, limited(config, max_file_bytes=len(before[0]))).rows) == 1
    with pytest.raises(IngestionError, match="max_file_bytes.*observed.*configured"):
        load_batch(tmp_path, limited(config, max_file_bytes=len(before[0]) - 1))
    assert (path.read_bytes(), path.stat().st_mtime_ns) == before


def test_file_read_is_bounded_even_if_metadata_understates_size(tmp_path, monkeypatch):
    path = tmp_path / "growing.csv"
    path.write_bytes(b"x" * 100)
    monkeypatch.setattr(ingest.os, "fstat", lambda fd: SimpleNamespace(st_size=0))
    with pytest.raises(LimitExceeded, match="observed 9, configured 8"):
        ingest._read_snapshot(path, InputLimits(max_file_bytes=8))


def test_csv_record_budget_counts_duplicates_invalid_and_blank_rows(
    tmp_path, write_csv, config_factory
):
    write_csv(tmp_path / "sales.csv", rows=[VALID, VALID, ["", *VALID[1:]], []])
    config = config_factory()
    assert len(load_batch(tmp_path, limited(config, max_total_records=4)).rows) == 4
    with pytest.raises(IngestionError, match="max_total_records.*observed 4, configured 3"):
        load_batch(tmp_path, limited(config, max_total_records=3))


@pytest.mark.parametrize("csv_first", [True, False])
def test_record_budget_is_batch_wide_for_mixed_formats(
    tmp_path, write_csv, write_xlsx, config_factory, profile_factory, csv_first
):
    write_csv(tmp_path / ("a.csv" if csv_first else "z.csv"))
    write_xlsx(tmp_path / ("z.xlsx" if csv_first else "a.xlsx"))
    config = config_factory(profile_factory(), profile_factory("xlsx"))
    assert len(load_batch(tmp_path, limited(config, max_total_records=2)).rows) == 2
    with pytest.raises(IngestionError, match="max_total_records.*observed 2, configured 1"):
        load_batch(tmp_path, limited(config, max_total_records=1))


def test_record_budget_spans_selected_sheets(tmp_path, write_xlsx, profile_factory, config_factory):
    write_xlsx(tmp_path / "sales.xlsx", sheets={"A": [VALID, VALID], "Z": [VALID]})
    config = config_factory(profile_factory("xlsx", sheets=("Z", "A")))
    assert len(load_batch(tmp_path, limited(config, max_total_records=3)).rows) == 3
    with pytest.raises(IngestionError, match="max_total_records.*observed 3, configured 2"):
        load_batch(tmp_path, limited(config, max_total_records=2))


@pytest.mark.parametrize(
    "option,maximum",
    [
        ("max_worksheet_rows", 1),
        ("max_worksheet_columns", 5),
    ],
)
def test_actual_worksheet_dimensions_are_bounded(xlsx_input, option, maximum):
    path, config = xlsx_input
    assert (
        len(
            load_batch(
                path.parent, limited(config, max_worksheet_rows=2, max_worksheet_columns=6)
            ).rows
        )
        == 1
    )
    with pytest.raises(IngestionError, match=option):
        load_batch(path.parent, limited(config, **{option: maximum}))


def test_stale_dimension_hints_do_not_override_actual_geometry(xlsx_input):
    path, config = xlsx_input
    with ZipFile(path) as archive:
        xml = archive.read(PART).replace(
            b'<dimension ref="A1:F2"', b'<dimension ref="A1:XFD1048576"'
        )
    rewrite_package(path, {PART: xml})
    assert (
        len(
            load_batch(
                path.parent, limited(config, max_worksheet_rows=2, max_worksheet_columns=6)
            ).rows
        )
        == 1
    )


def test_sparse_row_record_budget_fails_before_openpyxl_allocates_gaps(xlsx_input, monkeypatch):
    path, config = xlsx_input
    with ZipFile(path) as archive:
        xml = archive.read(PART).replace(b'<row r="2">', b'<row r="50">')
        xml = re.sub(rb'r="([A-F])2"', rb'r="\g<1>50"', xml)
    rewrite_package(path, {PART: xml})
    monkeypatch.setattr(
        ingest, "load_workbook", lambda *args, **kwargs: pytest.fail("decoded early")
    )
    with pytest.raises(IngestionError, match="max_total_records.*observed 49, configured 4"):
        load_batch(path.parent, limited(config, max_total_records=4))


@pytest.mark.parametrize(
    "option",
    [
        "max_zip_members",
        "max_zip_member_bytes",
        "max_zip_uncompressed_bytes",
        "max_zip_compression_ratio",
    ],
)
def test_zip_budget_boundaries(xlsx_input, option):
    path, config = xlsx_input
    with ZipFile(path) as archive:
        entries = archive.infolist()
    measured = {
        "max_zip_members": len(entries),
        "max_zip_member_bytes": max(item.file_size for item in entries),
        "max_zip_uncompressed_bytes": sum(item.file_size for item in entries),
        "max_zip_compression_ratio": max(
            (item.file_size + item.compress_size - 1) // item.compress_size for item in entries
        ),
    }[option]
    assert len(load_batch(path.parent, limited(config, **{option: measured})).rows) == 1
    with pytest.raises(IngestionError, match=option):
        load_batch(path.parent, limited(config, **{option: measured - 1}))


@pytest.mark.parametrize("option", ["max_zip_member_bytes", "max_zip_uncompressed_bytes"])
def test_inflation_checks_actual_bytes_not_only_zip_metadata(option):
    entry = SimpleNamespace(
        filename="data.bin", file_size=1, compress_size=1, flag_bits=0, compress_type=ZIP_STORED
    )
    archive = SimpleNamespace(infolist=lambda: [entry], open=lambda item: io.BytesIO(b"x" * 100))
    with pytest.raises(LimitExceeded, match=f"{option}.*observed 9, configured 8"):
        _checked_members(archive, replace(InputLimits(), **{option: 8}))


def test_excessive_compression_fails_before_any_inflation(xlsx_input, monkeypatch):
    path, config = xlsx_input
    rewrite_package(path, extra={"large.bin": b"0" * 16_384}, compression=ZIP_DEFLATED)
    monkeypatch.setattr(ZipFile, "open", lambda *args, **kwargs: pytest.fail("inflated early"))
    with pytest.raises(IngestionError, match="max_zip_compression_ratio"):
        load_batch(path.parent, config)


@pytest.mark.parametrize("name", ["../outside.xml", "/absolute.xml", "xl\\bad.xml", "C:/bad.xml"])
def test_unsafe_zip_member_paths_fail(xlsx_input, name):
    path, config = xlsx_input
    rewrite_package(path, extra={name: b"<x/>"})
    if "\\" in name:
        # The Windows ZIP writer normalizes separators. Prepare malformed local
        # and central headers explicitly, before exercising read-only ingestion.
        path.write_bytes(path.read_bytes().replace(name.replace("\\", "/").encode(), name.encode()))
    with pytest.raises(IngestionError, match="Unsafe ZIP member path"):
        load_batch(path.parent, config)


def test_crc_corruption_fails_explicitly(xlsx_input):
    path, config = xlsx_input
    payload = b"unique-crc-probe-data"
    rewrite_package(path, extra={"crc.bin": payload})
    contents = path.read_bytes()
    assert contents.count(payload) == 1
    path.write_bytes(contents.replace(payload, b"X" + payload[1:]))
    with pytest.raises(IngestionError, match="malformed ZIP member.*CRC"):
        load_batch(path.parent, config)


def test_xml_depth_and_element_boundaries():
    _parse_xml(b"<a><b/></a>", "depth.xml", InputLimits(max_xml_depth=2))
    with pytest.raises(LimitExceeded, match="max_xml_depth.*observed 2, configured 1"):
        _parse_xml(b"<a><b/></a>", "depth.xml", InputLimits(max_xml_depth=1))
    _parse_xml(b"<a><b/><b/></a>", "elements.xml", InputLimits(max_xml_elements=3))
    with pytest.raises(LimitExceeded, match="max_xml_elements.*observed 3, configured 2"):
        _parse_xml(b"<a><b/><b/></a>", "elements.xml", InputLimits(max_xml_elements=2))


@pytest.mark.parametrize("option", ["max_xml_depth", "max_xml_elements"])
def test_xml_budgets_cover_parts_not_selected_for_ingestion(xlsx_input, option):
    path, config = xlsx_input
    baseline = 0
    with ZipFile(path) as archive:
        for entry in archive.infolist():
            root = ET.fromstring(archive.read(entry))
            if option == "max_xml_elements":
                measured = sum(1 for _ in root.iter())
            else:
                stack = [(root, 1)]
                measured = 0
                while stack:
                    element, depth = stack.pop()
                    measured = max(measured, depth)
                    stack.extend((child, depth + 1) for child in element)
            baseline = max(baseline, measured)
    maximum = baseline + 1
    payload = (
        b"<a>" * (maximum + 1) + b"</a>" * (maximum + 1)
        if option == "max_xml_depth"
        else b"<a>" + b"<b/>" * maximum + b"</a>"
    )
    rewrite_package(path, extra={"customXml/extra.xml": payload})
    with pytest.raises(IngestionError, match=f"customXml/extra.xml: limits.{option}"):
        load_batch(path.parent, limited(config, **{option: maximum}))


@pytest.mark.parametrize(
    "encoding,name",
    [
        ("utf-8", "customXml/test.xml"),
        ("utf-16", "customXml/test.xml"),
        ("utf-16", "disguised.bin"),
    ],
)
def test_xml_dtd_entities_rejected_including_utf16_and_nonstandard_parts(
    xlsx_input, encoding, name
):
    path, config = xlsx_input
    payload = '<!DOCTYPE a [<!ENTITY text "expanded">]><a>&text;</a>'.encode(encoding)
    rewrite_package(path, extra={name: payload})
    with pytest.raises(IngestionError, match="DTD/entity declarations are not allowed"):
        load_batch(path.parent, config)


def test_external_dtd_is_rejected(xlsx_input):
    path, config = xlsx_input
    rewrite_package(
        path,
        extra={
            "customXml/extra.xml": b'<!DOCTYPE a SYSTEM "https://example.invalid/not-fetched.dtd"><a/>'
        },
    )
    with pytest.raises(IngestionError, match="DTD/entity declarations"):
        load_batch(path.parent, config)


def test_malformed_xml_is_rejected_even_in_unreferenced_parts(xlsx_input):
    path, config = xlsx_input
    rewrite_package(path, extra={"customXml/extra.xml": b"<a>"})
    with pytest.raises(IngestionError, match="customXml/extra.xml: malformed XML"):
        load_batch(path.parent, config)


def test_literal_declaration_text_is_not_mistaken_for_a_dtd(xlsx_input):
    path, config = xlsx_input
    rewrite_package(
        path,
        extra={
            "customXml/extra.xml": b"<a><!-- <!DOCTYPE a> --><b><![CDATA[<!ENTITY text>]]></b></a>"
        },
    )
    assert len(load_batch(path.parent, config).accepted) == 1


def test_cli_limit_failure_is_explicit_and_raising_limit_restores_processing(
    tmp_path, write_csv, write_config, capsys
):
    inputs = tmp_path / "inputs"
    path = write_csv(inputs / "sales.csv", rows=[VALID, VALID])
    config_path = write_config(tmp_path / "sources.toml")
    original_config = config_path.read_text()
    config_path.write_text(original_config + "\n[limits]\nmax_total_records=1\n")
    before = path.read_bytes(), path.stat().st_mtime_ns
    output = tmp_path / "out"
    args = ["--input", str(inputs), "--config", str(config_path), "--output", str(output)]
    assert main(args) == 2
    captured = capsys.readouterr()
    assert "max_total_records" in captured.err and "configured 1" in captured.err
    assert captured.out == "" and not output.exists()
    config_path.write_text(original_config + "\n[limits]\nmax_total_records=2\n")
    assert main(args) == 0
    capsys.readouterr()
    manifest = json.loads((output / "manifest.json").read_text())
    assert manifest["configuration"]["limits"]["max_total_records"] == 2
    assert manifest["configuration_sha256"] == load_config(config_path).fingerprint
    assert (path.read_bytes(), path.stat().st_mtime_ns) == before


def test_limit_changes_are_configuration_mismatches(tmp_path, write_csv, config_factory):
    inputs = tmp_path / "inputs"
    write_csv(inputs / "sales.csv")
    config = config_factory()
    result = consolidate(load_batch(inputs, config))
    with pytest.raises(ReconciliationError, match="Configuration mismatch"):
        export_batch(result, limited(config, max_total_records=3), tmp_path / "out")
    assert not (tmp_path / "out").exists()
