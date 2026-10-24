"""Independent conservation and policy checks before any master is published."""

from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal

from .config import Config
from .consolidate import ConsolidatedBatch, payload, source_order, transaction_key
from .models import RowResult, SourceRef
from .money import format_minor, minor_units

STATUSES = ("accepted", "rejected", "redundant")


class ReconciliationError(ValueError):
    """A broken conservation/policy invariant prevents publication."""


@dataclass(frozen=True)
class CurrencyReconciliation:
    currency: str
    precision: int
    counts: tuple[int, int, int, int]  # input, accepted, rejected, redundant
    totals: tuple[int, int, int, int]
    unquantified: tuple[int, int, int, int]

    def to_dict(self) -> dict:
        labels = ("input", *STATUSES)
        return {
            "currency": self.currency,
            "precision": self.precision,
            "known_amounts_balanced": True,
            "monetary_complete": self.unquantified[0] == 0,
            "counts": dict(zip(labels, self.counts, strict=True)),
            "amounts": {
                label: format_minor(total, self.precision)
                for label, total in zip(labels, self.totals, strict=True)
            },
            "minor_units": {
                label: str(total) for label, total in zip(labels, self.totals, strict=True)
            },
            "unquantified_counts": dict(zip(labels, self.unquantified, strict=True)),
        }


@dataclass(frozen=True)
class ReconciliationReport:
    counts: tuple[int, int, int, int]
    currencies: tuple[CurrencyReconciliation, ...]
    unknown_currency_count: int
    unquantified_count: int

    @property
    def complete(self) -> bool:
        return self.unquantified_count == 0

    def to_dict(self) -> dict:
        return {
            "counts": dict(zip(("input", *STATUSES), self.counts, strict=True)),
            "count_equation": "input = accepted + rejected + redundant",
            "count_balanced": True,
            "known_amounts_balanced": True,
            "monetary_complete": self.complete,
            "unknown_currency_count": self.unknown_currency_count,
            "unquantified_count": self.unquantified_count,
            "currencies": [currency.to_dict() for currency in self.currencies],
        }


def _check_money(row: RowResult, config: Config) -> None:
    currency = row.values.get("currency")
    amount = row.values.get("amount")
    expected = None
    if isinstance(amount, Decimal) and currency in config.currency_precision:
        try:
            expected = minor_units(amount, config.currency_precision[currency])
        except ValueError:
            pass  # Invalid precision/size is a legitimate unquantified rejection.
    if row.amount_minor != expected or (
        row.amount_minor is not None and type(row.amount_minor) is not int
    ):
        raise ReconciliationError(f"Incorrect minor-unit amount at {row.source}")
    if row.status == "accepted" and expected is None:
        raise ReconciliationError(f"Field-valid row has no quantifiable money: {row.source}")


def reconcile(result: ConsolidatedBatch, config: Config) -> ReconciliationReport:
    """Check source identity, partition, duplicate policy, and exact currency ledgers."""
    if result.batch.config_fingerprint != config.fingerprint:
        raise ReconciliationError(
            "Configuration mismatch: use the configuration applied at ingestion"
        )
    filenames = [file.source_file for file in result.batch.files]
    inventory = set(filenames)
    if len(inventory) != len(filenames):
        raise ReconciliationError("Input file inventory contains duplicate filenames")
    if any(row.source.source_file not in inventory for row in result.batch.rows):
        raise ReconciliationError("A record has no matching input file snapshot")
    originals = {row.source: row for row in result.batch.rows}
    if len(originals) != len(result.batch.rows):
        raise ReconciliationError("Input provenance is not unique")
    resolved = {item.row.source: item for item in result.rows}
    if len(resolved) != len(result.rows) or set(resolved) != set(originals):
        raise ReconciliationError("Input records were lost, added, or classified more than once")
    counts = Counter(item.status for item in result.rows)
    if set(counts) - set(STATUSES):
        raise ReconciliationError("Unknown terminal classification")
    count_tuple = (len(originals), *(counts[status] for status in STATUSES))
    if count_tuple[0] != sum(count_tuple[1:]):
        raise ReconciliationError("Count mismatch: input != accepted + rejected + redundant")

    accepted_keys: dict[tuple[str, str], SourceRef] = {}
    for item in result.rows:
        row = item.row
        if row != originals[row.source]:
            raise ReconciliationError(f"Classified record was changed: {row.source}")
        _check_money(row, config)
        if item.status == "accepted":
            key = transaction_key(row)
            if row.status != "accepted" or row.issues or item.issues or key is None:
                raise ReconciliationError("An invalid record was accepted")
            if key in accepted_keys:
                raise ReconciliationError("Master contains a duplicate business key")
            if item.duplicate_of is not None:
                raise ReconciliationError("Accepted record cannot reference a duplicate winner")
            accepted_keys[key] = row.source
        elif not item.issues:
            raise ReconciliationError("Excluded record has no reason")

    groups = defaultdict(list)
    for item in result.rows:
        key = transaction_key(item.row)
        if key is not None:
            groups[key].append(item)
        if item.status == "rejected":
            if key in accepted_keys:
                raise ReconciliationError("An accepted key also has a quarantined candidate")
            if item.duplicate_of is not None:
                raise ReconciliationError("Rejected records cannot be classified as duplicates")
        if item.status == "redundant":
            winner = resolved.get(item.duplicate_of)
            if (
                winner is None
                or winner.status != "accepted"
                or item.row.status != "accepted"
                or item.row.issues
                or payload(item.row) != payload(winner.row)
            ):
                raise ReconciliationError("Redundant record is not identical to an accepted winner")
    for candidates in groups.values():
        winners = [item for item in candidates if item.status == "accepted"]
        if winners and winners[0].row.source != min(
            (item.row.source for item in candidates), key=source_order
        ):
            raise ReconciliationError("Duplicate representative is not deterministic")
        if not winners:
            valid = all(item.row.status == "accepted" for item in candidates)
            identical = len({payload(item.row) for item in candidates}) == 1
            if valid and identical:
                raise ReconciliationError("Valid nonconflicting transaction was excluded")

    # Build the input ledger independently of the terminal-classification ledger.
    input_counts = Counter()
    input_totals = Counter()
    input_unknown = Counter()
    outcome_counts = Counter()
    outcome_totals = Counter()
    outcome_unknown = Counter()
    unknown_currency = 0
    unquantified = 0
    for row in result.batch.rows:
        currency = row.values.get("currency")
        if currency not in config.currency_precision:
            unknown_currency += 1
        if row.amount_minor is None:
            unquantified += 1
        if currency in config.currency_precision:
            input_counts[currency] += 1
            if row.amount_minor is None:
                input_unknown[currency] += 1
            else:
                input_totals[currency] += row.amount_minor
    for item in result.rows:
        row = item.row
        currency = row.values.get("currency")
        if currency in config.currency_precision:
            key = currency, item.status
            outcome_counts[key] += 1
            if row.amount_minor is None:
                outcome_unknown[key] += 1
            else:
                outcome_totals[key] += row.amount_minor
    currencies = []
    for currency, precision in sorted(config.currency_precision.items()):
        currency_counts = (
            input_counts[currency],
            *(outcome_counts[currency, status] for status in STATUSES),
        )
        totals = (
            input_totals[currency],
            *(outcome_totals[currency, status] for status in STATUSES),
        )
        unknown = (
            input_unknown[currency],
            *(outcome_unknown[currency, status] for status in STATUSES),
        )
        if any(values[0] != sum(values[1:]) for values in (currency_counts, totals, unknown)):
            raise ReconciliationError(f"Count/monetary reconciliation failed for {currency}")
        currencies.append(
            CurrencyReconciliation(currency, precision, currency_counts, totals, unknown)
        )
    return ReconciliationReport(count_tuple, tuple(currencies), unknown_currency, unquantified)
