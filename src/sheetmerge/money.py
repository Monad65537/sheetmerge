"""Exact money conversion and formatting, independent of Decimal context."""

from decimal import Decimal

MAX_MINOR_DIGITS = 38


def minor_units(amount: Decimal, precision: int) -> int:
    """Convert without rounding, accepting only exactly representable amounts."""
    if not amount.is_finite():
        raise ValueError("INVALID_AMOUNT")
    if amount.is_zero():
        return 0
    sign, digits, exponent = amount.as_tuple()
    shift = exponent + precision
    if shift < 0:
        remove = -shift
        if remove >= len(digits) or any(digits[len(digits) - remove :]):
            raise ValueError("EXCESS_FRACTIONAL_PRECISION")
        digits = digits[:-remove]
        shift = 0
    if len(digits) + shift > MAX_MINOR_DIGITS:
        raise ValueError("AMOUNT_TOO_LARGE")
    coefficient = int("".join(map(str, digits)))
    return (-1 if sign else 1) * coefficient * 10**shift


def format_minor(amount: int, precision: int) -> str:
    """Format arbitrarily large signed integer totals without float/Decimal math."""
    sign = "-" if amount < 0 else ""
    digits = str(abs(amount)).zfill(precision + 1)
    if precision == 0:
        return sign + digits
    return f"{sign}{digits[:-precision]}.{digits[-precision:]}"
