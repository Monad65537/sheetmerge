"""Explicit source profiles; no header, locale, or currency inference."""

import hashlib
import json
import re
import tomllib
from collections.abc import Mapping
from dataclasses import dataclass, field, fields
from pathlib import Path, PurePosixPath
from types import MappingProxyType

from .models import CANONICAL_COLUMNS


class ConfigurationError(ValueError):
    """Configuration cannot describe an input unambiguously."""


class LimitExceeded(ValueError):
    """A configured workload budget was exceeded; never truncate to fit."""


@dataclass(frozen=True)
class InputLimits:
    max_file_bytes: int = 20 * 1024 * 1024
    max_worksheet_rows: int = 100_001
    max_worksheet_columns: int = 256
    max_total_records: int = 100_000
    max_zip_members: int = 1_000
    max_zip_member_bytes: int = 32 * 1024 * 1024
    max_zip_uncompressed_bytes: int = 128 * 1024 * 1024
    max_zip_compression_ratio: int = 200
    max_xml_depth: int = 64
    max_xml_elements: int = 2_000_000

    def __post_init__(self):
        for option in fields(self):
            if type(getattr(self, option.name)) is not int or getattr(self, option.name) < 1:
                raise ConfigurationError(f"limits.{option.name} must be a positive integer")
        if self.max_worksheet_rows > 1_048_576 or self.max_worksheet_columns > 16_384:
            raise ConfigurationError("Worksheet limits cannot exceed Excel's physical dimensions")

    def check(self, name: str, observed: int, context: str) -> None:
        maximum = getattr(self, name)
        if observed > maximum:
            raise LimitExceeded(
                f"{context}: limits.{name} exceeded (observed {observed}, configured {maximum})"
            )


def _strings(value: object, label: str, *, required: bool = False) -> tuple[str, ...]:
    if not isinstance(value, (tuple, list)) or any(
        not isinstance(item, str) or not item for item in value
    ):
        raise ConfigurationError(f"{label} must be a list of nonempty strings")
    result = tuple(value)
    if required and not result:
        raise ConfigurationError(f"{label} cannot be empty")
    if len(set(result)) != len(result):
        raise ConfigurationError(f"{label} contains duplicates")
    return result


@dataclass(frozen=True)
class SourceProfile:
    name: str
    format: str
    patterns: tuple[str, ...]
    columns: Mapping[str, str]
    sheets: tuple[str, ...] = ()
    header_row: int = 1
    encoding: str = "utf-8-sig"
    delimiter: str = ","
    quotechar: str = '"'
    date_formats: tuple[str, ...] = ("%Y-%m-%d",)
    decimal_separator: str = "."
    grouping_separator: str | None = None
    trim_whitespace: bool = False
    uppercase_currency: bool = False
    allow_excel_dates: bool = False
    ignored_headers: tuple[str, ...] = ()

    def __post_init__(self):
        if not isinstance(self.name, str) or not self.name.strip():
            raise ConfigurationError("Profile name must be nonempty")
        if self.format not in ("csv", "xlsx"):
            raise ConfigurationError(f"{self.name}: format must be csv or xlsx")
        for name in ("patterns", "sheets", "date_formats", "ignored_headers"):
            object.__setattr__(
                self,
                name,
                _strings(getattr(self, name), name, required=name in ("patterns", "date_formats")),
            )
        for pattern in self.patterns:
            if (
                PurePosixPath(pattern).is_absolute()
                or ".." in PurePosixPath(pattern).parts
                or "\\" in pattern
                or ":" in pattern
            ):
                raise ConfigurationError("Patterns must be relative POSIX paths")
        if not isinstance(self.columns, Mapping) or set(self.columns) != set(CANONICAL_COLUMNS):
            raise ConfigurationError(f"{self.name}: map exactly all six canonical columns")
        headers = tuple(self.columns.values())
        if any(not isinstance(value, str) or not value.strip() for value in headers):
            raise ConfigurationError("Mapped source headers must be nonempty strings")
        if len(set(headers)) != len(headers):
            raise ConfigurationError("A source header cannot map to multiple canonical columns")
        if set(headers) & set(self.ignored_headers):
            raise ConfigurationError("Mapped headers cannot also be ignored")
        object.__setattr__(self, "columns", MappingProxyType(dict(self.columns)))
        if self.format == "xlsx" and not self.sheets:
            raise ConfigurationError("XLSX profiles must list included sheets explicitly")
        if self.format == "csv" and self.sheets:
            raise ConfigurationError("CSV profiles cannot specify sheets")
        if type(self.header_row) is not int or self.header_row < 1:
            raise ConfigurationError("header_row must be a positive integer")
        for name in ("trim_whitespace", "uppercase_currency", "allow_excel_dates"):
            if type(getattr(self, name)) is not bool:
                raise ConfigurationError(f"{name} must be a boolean")
        if not isinstance(self.encoding, str) or not self.encoding:
            raise ConfigurationError("encoding must be nonempty")
        for name in ("delimiter", "quotechar"):
            value = getattr(self, name)
            if not isinstance(value, str) or len(value) != 1 or value in "\r\n\0":
                raise ConfigurationError(f"{name} must be one non-newline character")
        if self.delimiter == self.quotechar:
            raise ConfigurationError("delimiter and quotechar must differ")
        if self.decimal_separator not in (".", ","):
            raise ConfigurationError("decimal_separator must be '.' or ','")
        grouping = self.grouping_separator
        if grouping is not None and (
            not isinstance(grouping, str)
            or len(grouping) != 1
            or grouping == self.decimal_separator
            or grouping.isalnum()
            or grouping in "+-\r\n\0"
        ):
            raise ConfigurationError("grouping_separator must be a distinct separator character")
        for date_format in self.date_formats:
            directives = re.findall(r"%.", date_format)
            if sorted(directives) != ["%Y", "%d", "%m"] or "%" in re.sub(
                r"%[Ymd]", "", date_format
            ):
                raise ConfigurationError("Date formats must contain only %Y, %m, %d, once each")


@dataclass(frozen=True)
class Config:
    currency_precision: Mapping[str, int]
    profiles: tuple[SourceProfile, ...]
    limits: InputLimits = field(default_factory=InputLimits)

    def __post_init__(self):
        if not isinstance(self.limits, InputLimits):
            raise ConfigurationError("limits must be an InputLimits instance")
        if not isinstance(self.currency_precision, Mapping) or not self.currency_precision:
            raise ConfigurationError("Configure at least one currency and its precision")
        for code, precision in self.currency_precision.items():
            if not isinstance(code, str) or not re.fullmatch(r"[A-Z]{3}", code):
                raise ConfigurationError("Currency codes must be three uppercase ASCII letters")
            if type(precision) is not int or not 0 <= precision <= 9:
                raise ConfigurationError("Currency precision must be an integer from 0 through 9")
        object.__setattr__(
            self, "currency_precision", MappingProxyType(dict(self.currency_precision))
        )
        if (
            not isinstance(self.profiles, (tuple, list))
            or not self.profiles
            or any(not isinstance(profile, SourceProfile) for profile in self.profiles)
        ):
            raise ConfigurationError("Configure at least one source profile")
        object.__setattr__(self, "profiles", tuple(self.profiles))
        if len({profile.name for profile in self.profiles}) != len(self.profiles):
            raise ConfigurationError("Source profile names must be unique")

    def profile_for(self, relative_path: str) -> SourceProfile:
        path = PurePosixPath(relative_path)
        matches = [
            profile
            for profile in self.profiles
            if path.suffix.lower() == f".{profile.format}"
            and any(path.match(pattern, case_sensitive=True) for pattern in profile.patterns)
        ]
        if len(matches) != 1:
            raise ConfigurationError(
                f"{relative_path}: expected exactly one source profile, found {len(matches)}"
            )
        return matches[0]

    @property
    def fingerprint(self) -> str:
        """Hash resolved settings, ignoring mapping and set-like option ordering."""
        profiles = []
        for profile in sorted(self.profiles, key=lambda item: item.name):
            options = {}
            for option in fields(profile):
                value = getattr(profile, option.name)
                options[option.name] = (
                    dict(value)
                    if isinstance(value, Mapping)
                    else sorted(value)
                    if isinstance(value, tuple)
                    else value
                )
            profiles.append(options)
        canonical = json.dumps(
            {
                "currency_precision": dict(self.currency_precision),
                "profiles": profiles,
                "limits": {
                    option.name: getattr(self.limits, option.name) for option in fields(self.limits)
                },
            },
            sort_keys=True,
            separators=(",", ":"),
            ensure_ascii=True,
        )
        return hashlib.sha256(canonical.encode("ascii")).hexdigest()


def load_config(path: str | Path) -> Config:
    """Read TOML and reject misspelled/unknown configuration options."""
    try:
        with Path(path).open("rb") as stream:
            data = tomllib.load(stream)
        if not {"currency_precision", "sources"} <= set(data) or set(data) - {
            "currency_precision",
            "sources",
            "limits",
        }:
            raise ConfigurationError(
                "Top-level keys must be currency_precision, sources, and optional limits"
            )
        sources = data["sources"]
        if not isinstance(sources, list):
            raise ConfigurationError("sources must be an array of source tables")
        return Config(
            data["currency_precision"],
            tuple(SourceProfile(**item) for item in sources),
            InputLimits(**data.get("limits", {})),
        )
    except ConfigurationError:
        raise
    except (OSError, ValueError, TypeError, LookupError) as exc:
        raise ConfigurationError(f"Cannot load configuration {path}: {exc}") from exc
