"""Ingestion results, retaining both original and normalized record values."""

from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from types import MappingProxyType
from typing import Literal

import pandas as pd

CANONICAL_COLUMNS = (
    "branch_id",
    "transaction_id",
    "transaction_date",
    "customer_id",
    "amount",
    "currency",
)
IDENTIFIER_COLUMNS = ("branch_id", "transaction_id", "customer_id")
PROVENANCE_COLUMNS = ("source_file", "source_sheet", "source_row")


@dataclass(frozen=True)
class SourceRef:
    source_file: str
    source_sheet: str
    source_row: int


@dataclass(frozen=True)
class Issue:
    code: str
    field: str
    message: str


@dataclass(frozen=True)
class CellValue:
    value: object
    kind: str = "text"
    numeric_text: str | None = None
    invalid_excel_date: bool = False
    original_text: str | None = None


@dataclass(frozen=True)
class RowResult:
    source: SourceRef
    raw_values: Mapping[str, object]
    values: Mapping[str, object]
    amount_minor: int | None
    issues: tuple[Issue, ...]
    status: Literal["accepted", "rejected", "blank"]
    raw_cells: Mapping[str, CellValue] = field(default_factory=dict)

    def __post_init__(self):
        object.__setattr__(self, "raw_values", MappingProxyType(dict(self.raw_values)))
        object.__setattr__(self, "values", MappingProxyType(dict(self.values)))
        object.__setattr__(self, "raw_cells", MappingProxyType(dict(self.raw_cells)))


@dataclass(frozen=True)
class InputFile:
    source_file: str
    sha256: str
    profile: str


@dataclass(frozen=True)
class BatchResult:
    input_root: Path
    files: tuple[InputFile, ...]
    rows: tuple[RowResult, ...]
    config_fingerprint: str

    @property
    def accepted(self) -> tuple[RowResult, ...]:
        return tuple(row for row in self.rows if row.status == "accepted")

    @property
    def rejected(self) -> tuple[RowResult, ...]:
        return tuple(row for row in self.rows if row.status == "rejected")

    @property
    def blank(self) -> tuple[RowResult, ...]:
        return tuple(row for row in self.rows if row.status == "blank")

    def to_frame(self) -> pd.DataFrame:
        """Return accepted records without pandas type or missing-value inference."""
        records = [
            {
                **row.values,
                "source_file": row.source.source_file,
                "source_sheet": row.source.source_sheet,
                "source_row": row.source.source_row,
                "amount_minor": row.amount_minor,
            }
            for row in self.accepted
        ]
        return pd.DataFrame(
            records,
            columns=[*CANONICAL_COLUMNS, *PROVENANCE_COLUMNS, "amount_minor"],
            dtype=object,
        )
