"""Deterministic composite-key resolution with full-group conflict quarantine."""

from collections import defaultdict
from dataclasses import dataclass
from typing import Literal

from .models import CANONICAL_COLUMNS, BatchResult, Issue, RowResult, SourceRef

BUSINESS_KEY = ("branch_id", "transaction_id")


def source_order(source: SourceRef) -> tuple[str, str, int]:
    return source.source_file, source.source_sheet, source.source_row


def transaction_key(row: RowResult) -> tuple[str, str] | None:
    values = tuple(row.values.get(field) for field in BUSINESS_KEY)
    if all(isinstance(value, str) and value.strip() for value in values):
        return values
    return None


def payload(row: RowResult) -> tuple[object, ...]:
    return tuple(row.values.get(field) for field in CANONICAL_COLUMNS)


@dataclass(frozen=True)
class ClassifiedRow:
    row: RowResult
    status: Literal["accepted", "rejected", "redundant"]
    issues: tuple[Issue, ...] = ()
    duplicate_of: SourceRef | None = None


@dataclass(frozen=True)
class ConsolidatedBatch:
    batch: BatchResult
    rows: tuple[ClassifiedRow, ...]

    @property
    def accepted(self) -> tuple[ClassifiedRow, ...]:
        return tuple(row for row in self.rows if row.status == "accepted")

    @property
    def rejected(self) -> tuple[ClassifiedRow, ...]:
        return tuple(row for row in self.rows if row.status == "rejected")

    @property
    def redundant(self) -> tuple[ClassifiedRow, ...]:
        return tuple(row for row in self.rows if row.status == "redundant")


def consolidate(batch: BatchResult) -> ConsolidatedBatch:
    """Retain every occurrence; never choose a winner from a conflicting group."""
    groups: dict[tuple[str, str], list[RowResult]] = defaultdict(list)
    classified = []
    for row in batch.rows:
        key = transaction_key(row)
        if key is None:
            if row.status == "blank":
                issues = (Issue("BLANK_ROW", "row", "Fully blank input record"),)
            else:
                issues = row.issues or (
                    Issue("MISSING_BUSINESS_KEY", "transaction", "Composite key is not readable"),
                )
            classified.append(ClassifiedRow(row, "rejected", issues))
        else:
            groups[key].append(row)

    for key in sorted(groups):
        candidates = sorted(groups[key], key=lambda row: source_order(row.source))
        invalid = any(row.status != "accepted" or row.issues for row in candidates)
        valid_payloads = {payload(row) for row in candidates if row.status == "accepted"}
        conflicting = len(valid_payloads) > 1
        if invalid or conflicting:
            group_issues = []
            if conflicting:
                fields = [
                    field
                    for field in CANONICAL_COLUMNS
                    if len(
                        {row.values.get(field) for row in candidates if row.status == "accepted"}
                    )
                    > 1
                ]
                group_issues.append(
                    Issue(
                        "CONFLICTING_TRANSACTION",
                        "transaction",
                        "Same-key transactions disagree on: " + ", ".join(fields),
                    )
                )
            if invalid and len(candidates) > 1:
                group_issues.append(
                    Issue(
                        "KEY_HAS_INVALID_CANDIDATE",
                        "transaction",
                        "Same-key group contains a record that failed field validation",
                    )
                )
            for row in candidates:
                classified.append(ClassifiedRow(row, "rejected", row.issues + tuple(group_issues)))
        else:
            representative = candidates[0]
            classified.append(ClassifiedRow(representative, "accepted"))
            for row in candidates[1:]:
                classified.append(
                    ClassifiedRow(
                        row,
                        "redundant",
                        (
                            Issue(
                                "EXACT_DUPLICATE", "transaction", "Identical normalized transaction"
                            ),
                        ),
                        representative.source,
                    )
                )

    classified.sort(
        key=lambda item: (transaction_key(item.row) or ("", ""), source_order(item.row.source))
    )
    return ConsolidatedBatch(batch, tuple(classified))
