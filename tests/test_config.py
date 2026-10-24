from dataclasses import replace

import pytest

from sheetmerge import Config, ConfigurationError, load_config


def test_load_toml_and_explicit_mapping(tmp_path):
    path = tmp_path / "sources.toml"
    path.write_text(
        """
[currency_precision]
PLN = 2
[[sources]]
name = "local"
format = "csv"
patterns = ["*.csv"]
delimiter = ";"
decimal_separator = ","
date_formats = ["%d/%m/%Y"]
[sources.columns]
branch_id = "Branch"
transaction_id = "Receipt"
transaction_date = "Day"
customer_id = "Customer"
amount = "Total"
currency = "Currency"
""",
        encoding="utf-8",
    )
    config = load_config(path)
    assert config.currency_precision == {"PLN": 2}
    assert config.profiles[0].columns["transaction_id"] == "Receipt"
    assert config.profiles[0].delimiter == ";"


@pytest.mark.parametrize(
    "overrides",
    [
        {"columns": {"branch_id": "Branch"}},
        {
            "columns": dict.fromkeys(
                [
                    "branch_id",
                    "transaction_id",
                    "transaction_date",
                    "customer_id",
                    "amount",
                    "currency",
                ],
                "same",
            )
        },
        {"header_row": 0},
        {"header_row": True},
        {"delimiter": "||"},
        {"delimiter": '"'},
        {"trim_whitespace": "yes"},
        {"date_formats": ("%m/%d",)},
        {"date_formats": ("%Y-%b-%d",)},
        {"decimal_separator": ":"},
        {"grouping_separator": "."},
        {"grouping_separator": "0"},
        {"patterns": ()},
        {"patterns": ("../*.csv",)},
        {"patterns": ("C:/sales/*.csv",)},
        {"ignored_headers": ("amount",)},
        {"sheets": ("Sales",)},
    ],
)
def test_invalid_source_profiles(profile_factory, overrides):
    with pytest.raises(ConfigurationError):
        profile_factory(**overrides)


def test_xlsx_sheets_must_be_explicit(profile_factory):
    with pytest.raises(ConfigurationError, match="included sheets"):
        profile_factory("xlsx", sheets=())


@pytest.mark.parametrize("currencies", [{}, {"usd": 2}, {"USD": True}, {"USD": -1}, {"USD": 10}])
def test_invalid_currency_configuration(profile_factory, currencies):
    with pytest.raises(ConfigurationError):
        Config(currencies, (profile_factory(),))


def test_profile_must_match_exactly_once(profile_factory, config_factory):
    profile = profile_factory()
    config = config_factory(profile, profile_factory(name="other"))
    with pytest.raises(ConfigurationError, match="found 2"):
        config.profile_for("sales.csv")
    with pytest.raises(ConfigurationError, match="found 0"):
        config_factory().profile_for("sales.xlsx")


def test_profiles_and_headers_are_case_sensitive(config_factory):
    with pytest.raises(ConfigurationError, match="found 0"):
        config_factory().profile_for("SALES.CSV")


def test_unknown_toml_option_is_not_ignored(tmp_path):
    path = tmp_path / "bad.toml"
    path.write_text('[currency_precision]\nUSD=2\n[[sources]]\nnaem="typo"\n')
    with pytest.raises(ConfigurationError):
        load_config(path)


def test_configuration_defensively_copies_mappings(profile_factory):
    columns = dict(profile_factory().columns)
    profile = profile_factory(columns=columns)
    columns["branch_id"] = "Changed"
    assert profile.columns["branch_id"] == "branch_id"
    with pytest.raises(TypeError):
        profile.columns["branch_id"] = "Changed"


def test_configuration_fingerprint_ignores_only_ordering(profile_factory):
    csv = profile_factory(
        patterns=("sales*.csv", "returns*.csv"),
        date_formats=("%Y-%m-%d", "%d/%m/%Y"),
        ignored_headers=("Note", "Extra"),
    )
    xlsx = profile_factory("xlsx", sheets=("Sales", "Returns"))
    original = Config({"USD": 2, "PLN": 2}, (csv, xlsx))
    reordered = Config(
        {"PLN": 2, "USD": 2},
        (
            replace(xlsx, sheets=tuple(reversed(xlsx.sheets))),
            replace(
                csv,
                patterns=tuple(reversed(csv.patterns)),
                date_formats=tuple(reversed(csv.date_formats)),
                ignored_headers=tuple(reversed(csv.ignored_headers)),
                columns=dict(reversed(list(csv.columns.items()))),
            ),
        ),
    )
    assert original.fingerprint == reordered.fingerprint
    assert len(original.fingerprint) == 64
    assert replace(original, currency_precision={"USD": 3, "PLN": 2}).fingerprint != (
        original.fingerprint
    )
