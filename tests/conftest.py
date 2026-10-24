import csv
import io
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook

from sheetmerge import Config, SourceProfile
from sheetmerge.models import CANONICAL_COLUMNS

HEADERS = list(CANONICAL_COLUMNS)
VALID = ["001", "000042", "2024-02-29", "000007", "12.34", "USD"]


@pytest.fixture
def profile_factory():
    def factory(format="csv", **overrides):
        options = {
            "name": format,
            "format": format,
            "patterns": (f"*.{format}",),
            "columns": dict(zip(HEADERS, HEADERS, strict=True)),
        }
        if format == "xlsx":
            options["sheets"] = ("Sales",)
        options.update(overrides)
        return SourceProfile(**options)

    return factory


@pytest.fixture
def config_factory(profile_factory):
    def factory(*profiles, currencies=None):
        return Config(
            {"USD": 2, "PLN": 2, "JPY": 0, "KWD": 3} if currencies is None else currencies,
            tuple(profiles) or (profile_factory(),),
        )

    return factory


@pytest.fixture
def write_csv():
    def write(path: Path, rows=None, headers=None, encoding="utf-8", **dialect):
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding=encoding, newline="") as stream:
            writer = csv.writer(stream, **dialect)
            writer.writerow(HEADERS if headers is None else headers)
            writer.writerows([VALID] if rows is None else rows)
        return path

    return write


@pytest.fixture
def write_xlsx():
    def write(path: Path, rows=None, headers=None, sheets=None):
        workbook = Workbook()
        workbook.remove(workbook.active)
        for name, data in (sheets or {"Sales": [VALID] if rows is None else rows}).items():
            sheet = workbook.create_sheet(name)
            sheet.append(HEADERS if headers is None else headers)
            for row in data:
                sheet.append(row)
        workbook.save(path)
        workbook.close()
        return path

    return write


@pytest.fixture
def rewrite_xlsx():
    """Alter only synthetic fixture XML to exercise stored numeric precision."""

    def rewrite(path, member, old, new):
        output = io.BytesIO()
        with ZipFile(io.BytesIO(path.read_bytes())) as source, ZipFile(output, "w") as target:
            for entry in source.infolist():
                content = source.read(entry.filename)
                if entry.filename == member:
                    assert old in content
                    content = content.replace(old, new)
                target.writestr(entry, content)
        path.write_bytes(output.getvalue())

    return rewrite


@pytest.fixture
def write_config():
    def write(path: Path):
        text = "[currency_precision]\nUSD=2\nPLN=2\nJPY=0\nKWD=3\n"
        for format in ("csv", "xlsx"):
            text += f'[[sources]]\nname="{format}"\nformat="{format}"\npatterns=["*.{format}"]\n'
            if format == "xlsx":
                text += 'sheets=["Sales"]\nallow_excel_dates=true\n'
            text += "[sources.columns]\n"
            text += "\n".join(f'{field}="{field}"' for field in HEADERS) + "\n"
        path.write_text(text, encoding="utf-8")
        return path

    return write
