# SheetMerge

SheetMerge consolidates heterogeneous sales CSV and XLSX files into an auditable
master dataset. It applies explicit schema mappings, validates every record,
resolves composite-key duplicates, quarantines conflicts, reconciles exact monetary
totals per currency, and exports CSV and Excel alongside exclusion and audit reports.

Inputs are opened read-only and never modified. Every ingested data record receives
exactly one final outcome: `accepted`, `rejected`, or `redundant`.

## Setup and verification

Requires Python 3.12 or newer. Runtime dependencies are pandas and openpyxl;
development dependencies are pytest and Ruff.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
python -m pytest -q
python -m ruff check .
```

Tests generate synthetic fixtures in pytest temporary directories. If the system
temporary directory is inaccessible in a sandbox, set `TMP` and `TEMP` to an
accessible temporary directory before running pytest.

## Run a batch

For a ready-made mixed CSV/XLSX example, see the [12 synthetic sample files and
their branch mappings](samples/README.md).

Configure the exact source headers first, then run:

```powershell
python -m sheetmerge --config config/sources.example.toml --input inputs --output outputs/run-001
```

The installed `sheetmerge` command accepts the same arguments. `--help` describes
the arguments; `--version` prints the version. The output directory must be new
and outside the input tree. Existing outputs are never overwritten. Each run reads
the complete supplied batch; previous masters are not automatically merged back in.

The command prints a JSON summary on stdout. Errors go to stderr.

| Exit code | Meaning |
| --- | --- |
| `0` | Completed with no rejected rows and complete monetary coverage; reported exact duplicates are allowed |
| `1` | Output published for review: rejected/quarantined rows or incomplete monetary coverage |
| `2` | Configuration, input, reconciliation, or export failure; no completed bundle published |

A conflict-only batch produces an empty master plus explicit exclusions and exit
code `1`. Invalid amounts/currencies never produce a clean success status.

Each completed output directory contains:

| Artifact | Contents |
| --- | --- |
| `master.csv` | Accepted canonical values and representative source provenance |
| `master.xlsx` | The same values in the `Master` worksheet, with filters and a frozen header |
| `exclusions.csv` | Every rejected and redundant occurrence, structured reasons, raw values, and duplicate references |
| `audit.jsonl` | Every input record, its outcome, raw cell metadata, normalized values, and reasons |
| `reconciliation.csv` | Per-currency counts, known monetary totals, and unquantified counts |
| `reconciliation.json` | Global count equation, monetary coverage, unknown-currency count, and currency ledgers |
| `manifest.json` | Source hashes, configuration, policies, tool version, reconciliation, and artifact hashes |

Master columns are the six canonical fields followed by `source_file`,
`source_sheet`, and `source_row`. Dates use ISO calendar notation; amounts use the
configured currency's fixed precision. Excel stores every master cell as text,
including amounts, so long identifiers and large monetary values remain exact.
Formula-like IDs remain literal strings in the workbook.

CSV is UTF-8, fully quoted, with LF record terminators. Use `master.xlsx` for opening
directly in Excel. CSV quoting does not prevent application type/formula inference;
import CSV columns as text. In pandas, use `dtype=str, keep_default_na=False`.

Both masters are read back and compared with the reconciled accepted records.
Files are staged together, then published by a directory rename. A failed writer,
failed read-back, or failed reconciliation does not publish a partial bundle. Excel
row/cell limits and unsupported XML characters cause explicit export failure,
rather than truncation or omission.

## Configure each source explicitly

Start with [config/sources.example.toml](config/sources.example.toml). Mappings run
from canonical names to **exact source header names**:

```toml
[currency_precision]
USD = 2

[[sources]]
name = "branch-sales"
format = "csv"
patterns = ["*.csv"]
date_formats = ["%Y-%m-%d"]

[sources.columns]
branch_id = "Branch"
transaction_id = "Receipt"
transaction_date = "Date"
customer_id = "Customer"
amount = "Total"
currency = "Currency"
```

All six mappings are required. A supported input must match exactly one profile;
unknown files and overlapping matches fail the batch. Profile names must be unique.
Patterns are case-sensitive relative POSIX paths, using `PurePosixPath.match`:
`*.csv` matches basenames at any depth; `branch/*.csv` restricts the path suffix.
`**` does not provide recursive-glob semantics in this matcher. Use separate
profiles or explicit patterns where different branches have different schemas.

Optional source settings:

| Setting | Default / behavior |
| --- | --- |
| `header_row` | `1`; one-based physical row or CSV record's starting line |
| `encoding` | `utf-8-sig`; strict CSV decoding, no replacement characters |
| `delimiter`, `quotechar` | `,` and `"`; no dialect detection |
| `sheets` | Required explicit worksheet names for XLSX; other sheets excluded |
| `date_formats` | `["%Y-%m-%d"]`; each format uses `%Y`, `%m`, `%d` once |
| `decimal_separator` | `.`; may instead be `,` |
| `grouping_separator` | Disabled; optional distinct character with groups of three |
| `trim_whitespace` | `false`; explicit opt-in for all mapped string values |
| `uppercase_currency` | `false`; explicit opt-in |
| `allow_excel_dates` | `false`; opt in to native Excel dates |
| `ignored_headers` | `[]`; extra columns must be explicitly listed here |

Source headers are never trimmed or inferred. Missing, repeated, empty, or
unexpected headers fail the batch. Ignored columns are still retained as raw
audit values. Configured date formats must match exactly, including zero-padding;
a date with multiple different valid interpretations is rejected.

## Duplicate and conflict policies

The fixed composite key is **`(branch_id, transaction_id)`**. Each source row must
represent one complete transaction, with transaction IDs unique within a branch
across the supplied batch. Line-item files need an explicitly agreed transformation
before ingestion; a date is not automatically added to the key.

- A unique valid transaction is accepted.
- Identical normalized values across all six canonical fields retain one accepted
  representative. Every other occurrence is `redundant`, with `EXACT_DUPLICATE` and
  a reference to the representative's file, sheet, and row.
- Same-key valid transactions that disagree on amount, currency, customer, or date
  are all rejected with `CONFLICTING_TRANSACTION`. Even exact repeats within that
  group are quarantined. There is no latest-file or majority winner.
- An invalid record with a readable key blocks all other candidates sharing that
  key, with `KEY_HAS_INVALID_CANDIDATE`. Original field-level reasons are retained.
- Missing/unreadable keys are rejected independently; numeric IDs and lost zeroes
  are never inferred to associate them with another transaction.
- Fully blank data records become `rejected` with `BLANK_ROW`, preserving the count
  equation rather than disappearing from the consolidated result.

Duplicates are detected within files, across files, and across selected worksheets.
Raw ignored-column differences do not change transaction equality. Exact amounts
such as `1.2300` and `1.23` compare equal after validation. Corrections that reuse a
key with different values require upstream review and a fresh batch.

Representatives use ascending relative filename, sheet name, and physical row.
Master rows use ascending branch/transaction ID text, then provenance; other reports
use the same ordering, placing unreadable keys first. Sorting preserves case and
zeroes and uses Unicode lexical order rather than numeric/natural sorting. Currency
reports sort by currency code. XLSX core and ZIP timestamps are fixed to make the
complete artifact bundle byte-reproducible for identical input bytes/configuration
under the same runtime and library versions.

## Reconciliation contract

Every ingested data record, including blank records, must satisfy:

```text
input count = accepted count + rejected count + redundant count
```

Configured preambles and headers are metadata and are outside this equation.
Reconciliation checks the complete provenance partition, unchanged input records,
unique accepted keys, exclusion reasons, valid duplicate references, deterministic
representatives, and the absence of accepted candidates from quarantined keys.

For each configured currency, using exact Python integer minor units:

```text
input known amount = accepted amount + rejected known amount + redundant amount
```

Rejected records with a parseable amount/currency still contribute their exact
amount—for example, a row rejected for an invalid date. Integer totals may grow
beyond the per-transaction size limit without losing precision. Currencies are
never added together into a monetary grand total.

Invalid/unrepresentable amounts and unknown currencies are explicitly counted as
unquantified. Their amounts are not substituted with zero. Per-currency CSV columns
are named `*_known_amount`; unquantified counts and `monetary_complete` state the
coverage limits. The JSON report also accounts for records without a known currency.
Balanced known amounts do not imply complete financial coverage, and monetary
coverage does not establish agreement with external business totals.

A failed conservation or duplicate-policy invariant raises `ReconciliationError`
and prevents publication. Incomplete monetary coverage produces review outputs,
`monetary_complete=false`, and exit code `1`. Rejected conflicts with fully parseable
money can have complete monetary coverage while still requiring review.

## Configurable input limits

An optional top-level `[limits]` table controls workload budgets. Existing configs
without the table use the defaults below; individual values may be overridden.
These are starting defaults for a small local batch workload, **not universal safe
thresholds**. Tune them together for expected source files, measured memory use,
and the machine running the tool. A byte or record limit does not bound process
memory, CPU time, or every resource consumed by Python, pandas, and openpyxl.

```toml
[limits]
max_file_bytes = 20_971_520                 # 20 MiB per CSV/XLSX input
max_worksheet_rows = 100_001               # Highest actual worksheet row index
max_worksheet_columns = 256                # Highest actual worksheet column index
max_total_records = 100_000                # Across the complete input batch
max_zip_members = 1_000                    # Per XLSX archive, including directories
max_zip_member_bytes = 33_554_432           # 32 MiB uncompressed per ZIP member
max_zip_uncompressed_bytes = 134_217_728    # 128 MiB uncompressed across one archive
max_zip_compression_ratio = 200            # Maximum member uncompressed/compressed ratio
max_xml_depth = 64                        # Element nesting depth per XML document
max_xml_elements = 2_000_000                # Element count per XML document
```

Values must be positive integers; zero does not disable a limit. Worksheet settings
cannot exceed Excel's physical maximum of 1,048,576 rows and 16,384 columns. Limits
are inclusive: exactly the configured maximum is allowed. Failures identify the
limit, observed value, and configured value; compression ratios are reported rounded
up to the next integer. Exceeding a limit fails the entire batch with CLI exit code
`2` and no completed output bundle. Records are never truncated to fit. Raise the
specific budget only after checking the workload; existing files remain untouched.

The file-size check uses the opened file's size and a bounded read, so a growing
file cannot bypass the read limit. The record budget includes accepted, rejected,
redundant, and blank data records across CSVs and selected XLSX sheets. Preambles
and headers are excluded. A multiline CSV record counts once. Worksheet gaps that
the loader represents as blank records count too, and their spans are checked
before the decoder allocates them. The default 100,001-row setting allows a row-1
header plus 100,000 records; other header positions consume some of that span.

Worksheet dimensions mean actual row/cell coordinates, including formatted empty
rows/cells. Stale `<dimension>` hints remain ignored. Worksheet geometry and record
limits apply to selected sheets. ZIP expansion and XML checks apply to the whole
package, including unused XML parts, styles, and shared strings, because workbook
decoding may read more than the selected transaction cells.

XLSX preflight checks advertised sizes before inflation, then reads members in
bounded chunks, checks actual byte totals, and verifies sizes and CRCs. It rejects
duplicate/unsafe member paths, encrypted members, and compression methods other
than ZIP Stored or Deflate. XML parsing enforces depth/element budgets, rejects
malformed `.xml`/`.rels` parts, and forbids DTD/entity declarations regardless of
their encoding. Nonstandard part names are scanned for XML too; binary parts need
not be XML. These checks do not extract files or evaluate workbook formulas.

Resolved limits are part of the batch configuration fingerprint and manifest.
Exporting a previously ingested batch with different limits is a configuration
mismatch. Python callers can use `InputLimits` through `Config(..., limits=...)`.
Processing remains in memory; this is not a complete OOXML schema validator or a
sandbox for arbitrary hostile files.

## Python API

```python
from sheetmerge import consolidate, export_batch, load_batch, load_config, reconcile

config = load_config("config/sources.example.toml")
batch = load_batch("inputs", config)
result = consolidate(batch)
report = reconcile(result, config)
exported = export_batch(result, config, "outputs/run-001")  # Rechecks reconciliation

for item in result.rejected:
    print(item.row.source, [(issue.field, issue.code) for issue in item.issues])
```

`discover_inputs(path)` recursively lists `.csv` and `.xlsx` files in deterministic
relative-path order. Discovery recognizes extension case, but profile matching is
case-sensitive. Use a dedicated input directory: every discovered supported file
must have a profile. Symbolic links and Windows junctions are rejected.

`batch.files` records each relative filename, source profile, and SHA-256 of the
exact byte snapshot parsed. Files are opened read-only and are never saved or
modified. The caller should keep sources stable while a snapshot is being read.
Files and decoded rows are processed in memory within the configured input budgets;
large-scale streaming is outside this version's scope.

Selected XLSX sheets are structurally validated before value decoding. Duplicate
or out-of-order rows/cells, inconsistent cell-to-row coordinates, out-of-bounds
coordinates, duplicate ZIP parts, and ambiguous cell payloads fail the batch.
Sparse rows and unambiguous OOXML-omitted row/cell coordinates remain supported.
Monetary XML text and decoded values come from the same validated coordinate
structure. Stale worksheet dimension hints are still ignored. This is a focused
validator for supported worksheets, not a complete OOXML schema validator.

`batch.config_fingerprint` binds the applied configuration to the batch. Export and
reconciliation reject a different fingerprint before writing output. The canonical
SHA-256 includes all resolved settings and mappings; mapping order, profile order,
and set-like option order do not affect it. The manifest records it as
`configuration_sha256`. Equivalent reloaded configurations are allowed.

`batch.rows` includes each ingested record as field-valid `accepted`, `rejected`, or
`blank`. These are ingestion statuses; only `consolidate(batch)` resolves duplicates
and conflicts into the final `accepted`, `rejected`, and `redundant` outcomes.
`result.rows` contains `ClassifiedRow` objects retaining the original row, terminal
status, reasons, and optional `duplicate_of` reference. File order, worksheet-name
order, physical row order, and issue order are deterministic. Configured preambles
and headers are metadata rather than data records. A header-only source is valid;
an input directory containing no supported files is an error.

Each row contains:

- `source`: relative `source_file`, `source_sheet` (empty for CSV), and `source_row`.
  CSV rows use the first physical line, even for quoted multiline records.
- `raw_values`: original source columns; numeric Excel values and native Excel
  dates retain their original XML value text, including serials or ISO strings.
- `raw_cells`: extracted cell values, cell types, stored numeric text, and
  `original_text` for numeric/native date cells. Any float exposed here is audit
  metadata, never a monetary operand.
- `values`: canonical normalized fields, with `None` where conversion failed.
  Parseable amounts remain available even if another field failed validation.
- `amount_minor`: exact Python integer in the configured currency's minor units,
  or `None` if amount/currency validation prevents conversion.
- `issues`: all independent rejection reasons, each with code, field, and message.

`batch.to_frame()` is an ingestion inspection helper returning field-valid rows
before key resolution, with the six canonical columns, provenance, and `amount_minor`.
It must not be used as the final master. `export_batch` requires a consolidated
result and rechecks all invariants. Mapping containers are immutable; inspection
DataFrames are independent mutable views of values.

## Validation and monetary rules

Identifiers must be nonempty strings. Leading zeroes, long digit sequences, case,
Unicode, and punctuation are preserved. Literal `NA` is an identifier, not a
missing-value marker. Surrounding whitespace is preserved unless trimming is
explicitly enabled; whitespace-only fields are missing. Excel numeric identifier
cells are rejected even when their display format contains zero-padding. Missing
zeroes are never reconstructed.

Every canonical field is required. Dates must be valid calendar dates with no
time component. Native Excel dates require opt-in; the fictitious Excel date
1900-02-29 is rejected. Formulas and Excel error cells in mapped columns are
rejected rather than evaluated or replaced with cached values.

For native ISO Excel date cells (`t="d"`), the exact accepted forms are
`YYYY-MM-DD` and `YYYY-MM-DDT00:00:00`, optionally followed by a decimal point and
one or more zero fractional-second digits. Nonzero times/fractions, timezones,
whitespace, partial timestamps, invalid calendar dates, and trailing text are
rejected with `INVALID_DATE`; original text remains in exclusions and audit output.
Numeric Excel date serials must represent an integral calendar day, use the
workbook's date epoch, and decode to a valid date. CSV and ordinary Excel text-date
cells continue to use only the explicitly configured `date_formats`.

Currencies are explicitly configured three-letter uppercase codes, with precision
from zero through nine. Zero and negative amounts are permitted. Monetary strings
must match configured separators; scientific notation, currency symbols, NaN,
infinity, and accounting parentheses are rejected. Numeric Excel cells may use
scientific notation in their original XML representation.

Amounts are parsed as exact `Decimal` values from text, then converted to Python
integer minor units without rounding or Decimal-context arithmetic. Numeric Excel
amounts use stored XML decimals, never openpyxl's float conversion. No tool can
recover digits already lost before a workbook was saved. Extra trailing fractional
zeroes are accepted only when the amount is exactly representable at the configured
precision. Nonzero minor-unit amounts are limited to 38 digits; amount tokens are
limited to 512 characters. Use `amount_minor` for future aggregation rather than
summing Decimal objects under an arbitrary precision context.

Stable row reason codes include `MISSING_REQUIRED_VALUE`, `ID_NOT_TEXT`,
`INVALID_DATE`, `INVALID_AMOUNT`, `UNSUPPORTED_CURRENCY`,
`EXCESS_FRACTIONAL_PRECISION`, `AMOUNT_TOO_LARGE`, `FORMULA_NOT_ALLOWED`, and
`EXCEL_CELL_ERROR`. A row may have several reasons. Fully blank records are
accounted for separately; partially populated records are validated normally.

`ConfigurationError` reports invalid mappings/profiles. `IngestionError` reports
unreadable inputs, malformed CSV structure/field counts, invalid workbook
structure, missing selected sheets, or mismatched headers. These errors prevent a
partial batch result. Row-level validation failures remain inspectable in a
successful ingestion result. `ExportError` reports output collisions, unsupported
Excel content/limits, write failures, or master read-back differences.

## Scope and tests

This is a local, in-memory batch tool. There is no incremental database, currency
conversion, automatic source mapping, formula evaluation, automatic correction
selection, or source editing. Business key and currency precision policies should
be checked against the company's actual transaction conventions before use.

Tests cover source parsing and normalization, all duplicate/conflict outcomes,
invalid same-key candidates, count conservation, exact per-currency totals,
unquantified values, corrupted classification detection, deterministic artifacts,
literal Excel text, embedded newlines, input immutability, export limits, staged
failure cleanup, read-back verification, and CLI exit statuses. All fixtures are
synthetic.

