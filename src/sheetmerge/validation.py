"""Strict row validation with exact, context-independent monetary conversion."""

import re
from collections.abc import Mapping
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation

from .config import Config, SourceProfile
from .models import CANONICAL_COLUMNS, IDENTIFIER_COLUMNS, CellValue, Issue, RowResult, SourceRef
from .money import minor_units


def _amount(value: object, cell: CellValue, profile: SourceProfile) -> Decimal:
    if cell.kind == "numeric":
        # openpyxl may have already converted the value to a float. Never use that
        # float for money; use the original <v> text from the workbook XML instead.
        text = cell.numeric_text
        pattern = r"[+-]?[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?"
        if text is None or len(text) > 512 or not re.fullmatch(pattern, text):
            raise ValueError("Numeric Excel amount has no supported stored decimal value")
    elif cell.kind == "text" and isinstance(value, str):
        grouping = profile.grouping_separator
        integer = r"[0-9]+"
        if grouping is not None:
            integer = rf"(?:[0-9]+|[0-9]{{1,3}}(?:{re.escape(grouping)}[0-9]{{3}})+)"
        pattern = rf"[+-]?{integer}(?:{re.escape(profile.decimal_separator)}[0-9]+)?"
        if len(value) > 512 or not re.fullmatch(pattern, value):
            raise ValueError("Amount does not match the configured decimal/grouping syntax")
        text = value.replace(grouping, "") if grouping is not None else value
        text = text.replace(profile.decimal_separator, ".")
    else:
        raise ValueError("Amount must be decimal text or a numeric Excel cell")
    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("Amount is not a finite decimal") from exc
    if not amount.is_finite():
        raise ValueError("Amount is not a finite decimal")
    return amount


def _date(value: object, cell: CellValue, profile: SourceProfile) -> date:
    if cell.invalid_excel_date:
        raise ValueError("Excel's fictitious 1900-02-29 is not a valid date")
    if cell.kind == "date" and isinstance(value, str):
        if not profile.allow_excel_dates:
            raise ValueError("Native Excel dates require allow_excel_dates=true")
        original = cell.original_text if cell.original_text is not None else value
        if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}(?:T00:00:00(?:\.0+)?)?", original):
            raise ValueError(
                "Native ISO date must be YYYY-MM-DD or midnight YYYY-MM-DDT00:00:00; "
                "only zero fractional seconds are allowed, without timezone or extra text"
            )
        return date.fromisoformat(original[:10])
    if isinstance(value, date) and cell.numeric_text is not None:
        serial = Decimal(cell.numeric_text)
        if serial != serial.to_integral_value():
            raise ValueError(
                "Transaction date must not contain a time, including submillisecond time"
            )
    if isinstance(value, datetime):
        if not profile.allow_excel_dates:
            raise ValueError("Native Excel dates require allow_excel_dates=true")
        if value.tzinfo is not None or value.time() != time():
            raise ValueError("Transaction date must not contain a time or timezone")
        return value.date()
    if isinstance(value, date):
        if profile.allow_excel_dates:
            return value
        raise ValueError("Native Excel dates require allow_excel_dates=true")
    if not isinstance(value, str):
        raise ValueError("Date must be text in a configured format or an allowed Excel date")
    matches = set()
    for date_format in profile.date_formats:
        try:
            parsed = datetime.strptime(value, date_format).date()
        except ValueError:
            continue
        if parsed.strftime(date_format) == value:
            matches.add(parsed)
    if len(matches) > 1:
        raise ValueError("Date matches configured formats with different interpretations")
    if not matches:
        raise ValueError("Date does not match a configured format or is not a valid date")
    return matches.pop()


def validate_row(
    cells: Mapping[str, CellValue], source: SourceRef, profile: SourceProfile, config: Config
) -> RowResult:
    """Classify every data row, accumulating all independent field failures."""
    raw = {
        header: (
            cell.original_text
            if cell.original_text is not None
            else (
                cell.numeric_text
                if cell.kind == "numeric" and cell.numeric_text is not None
                else cell.value
            )
        )
        for header, cell in cells.items()
    }
    if all(cell.value is None or cell.value == "" for cell in cells.values()):
        return RowResult(source, raw, {}, None, (), "blank", cells)

    values: dict[str, object] = dict.fromkeys(CANONICAL_COLUMNS)
    issues: list[Issue] = []
    available: dict[str, tuple[object, CellValue]] = {}
    for field in CANONICAL_COLUMNS:
        cell = cells[profile.columns[field]]
        value = cell.value
        if cell.kind in ("formula", "error"):
            code = "FORMULA_NOT_ALLOWED" if cell.kind == "formula" else "EXCEL_CELL_ERROR"
            issues.append(Issue(code, field, "Mapped formula/error cells cannot be trusted"))
            continue
        if value is None or isinstance(value, str) and not value.strip():
            issues.append(Issue("MISSING_REQUIRED_VALUE", field, "Required value is empty"))
            continue
        if isinstance(value, str) and profile.trim_whitespace:
            value = value.strip()
        available[field] = value, cell

    for field in IDENTIFIER_COLUMNS:
        if field not in available:
            continue
        value, _ = available[field]
        if not isinstance(value, str) or available[field][1].kind != "text":
            issues.append(Issue("ID_NOT_TEXT", field, "Identifiers must be stored as text"))
        else:
            values[field] = value

    if "currency" in available:
        value, cell = available["currency"]
        if isinstance(value, str) and profile.uppercase_currency:
            value = value.upper()
        if (
            cell.kind != "text"
            or not isinstance(value, str)
            or value not in config.currency_precision
        ):
            issues.append(Issue("UNSUPPORTED_CURRENCY", "currency", "Currency is not configured"))
        else:
            values["currency"] = value

    if "transaction_date" in available:
        value, cell = available["transaction_date"]
        try:
            values["transaction_date"] = _date(value, cell, profile)
        except ValueError as exc:
            issues.append(Issue("INVALID_DATE", "transaction_date", str(exc)))

    amount_minor = None
    if "amount" in available:
        value, cell = available["amount"]
        try:
            amount = _amount(value, cell, profile)
            values["amount"] = amount
        except ValueError as exc:
            issues.append(Issue("INVALID_AMOUNT", "amount", str(exc)))
        else:
            currency = values["currency"]
            if currency is not None:
                try:
                    amount_minor = minor_units(amount, config.currency_precision[currency])
                except ValueError as exc:
                    issues.append(Issue(str(exc), "amount", "Amount cannot be represented exactly"))

    return RowResult(
        source,
        raw,
        values,
        amount_minor,
        tuple(issues),
        "rejected" if issues else "accepted",
        cells,
    )
