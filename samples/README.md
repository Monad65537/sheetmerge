# Synthetic sales fixtures

All identifiers, transactions, and amounts are invented. There are exactly 12 input
files: six CSVs and six XLSX workbooks, with three data records in each. No real
customer or business data is included. IDs such as `001`, `000001`, and `000101`
are deliberately stored as strings, including in Excel.

The [source configuration](sources.toml) has six explicit profiles: one per branch
and format. Every workbook contains a single `Sales` worksheet, with headers on
row 1 and data on physical rows 2–4. Excel dates are native dates and amounts are
numeric cells; CSVs exercise the branch-specific text formats below.

| Branch | ID / currency | Source headers | CSV conventions |
| --- | --- | --- | --- |
| A | `001` / USD | `Branch ID`, `Transaction ID`, `Date`, `Customer ID`, `Amount`, `Currency` | Comma delimiter, `YYYY-MM-DD`, decimal point |
| B | `002` / PLN | `store`, `sale_id`, `sold_at`, `client_ref`, `net_value`, `ccy` | Semicolon delimiter, `DD/MM/YYYY`, decimal comma |
| C | `003` / JPY | `location_code`, `order_no`, `order_date`, `account_id`, `sales_total`, `currency_code` | Comma delimiter, `YYYY/MM/DD`, whole-unit amounts |

## File inventory

Paths below are relative to `samples/input/`.

| File | Deliberate cases |
| --- | --- |
| `processable/branch_a/branch_a_01_sales.csv` | Leading-zero IDs; baseline transactions; USD 0.01 |
| `processable/branch_a/branch_a_02_duplicates_conflicts.xlsx` | Repeats A `000001`; A `000002` conflicts on amount: 50.00 versus 55.00 |
| `processable/branch_a/branch_a_03_invalid.csv` | A `000005` has a blank amount; A `000006` has invalid date `2024-02-30`; valid zero amount |
| `processable/branch_a/branch_a_04_refunds.xlsx` | Valid negative refund and additional valid transactions |
| `processable/branch_b/branch_b_01_sales.csv` | Leading-zero IDs; decimal commas and semicolons; baseline transactions |
| `processable/branch_b/branch_b_02_duplicates_conflicts.xlsx` | Repeats B `000001`; B `000002` conflicts on customer: `000202` versus `000299` |
| `processable/branch_b/branch_b_03_invalid.csv` | B `000005` has a blank amount; B `000006` has invalid date `31/04/2024`; valid zero amount |
| `processable/branch_b/branch_b_04_internal_duplicate.xlsx` | Valid negative refund; two identical B `000009` candidates have excess stored monetary precision and are quarantined |
| `processable/branch_c/branch_c_01_sales.csv` | Leading-zero IDs; whole-unit JPY amounts; baseline transactions |
| `processable/branch_c/branch_c_02_duplicates_conflicts.xlsx` | Repeats C `000001`; C `000002` conflicts on date: January 3 versus January 4 |
| `processable/branch_c/branch_c_03_invalid.csv` | C `000005` has a blank amount; C `000006` has invalid date `2024/13/01`; valid negative refund |
| `structural_errors/branch_c_04_missing_columns.xlsx` | Missing `account_id` and `currency_code` headers; intentionally blocks the full batch |

Repeated transaction numbers across different branches are legitimate separate
composite keys. Exact duplicates compare equal across CSV and Excel despite their
different date/amount representations. The internal B `000009` pair intentionally
retains a serialization artifact: both numeric cells store `99.98999999999999`
while displaying `99.99`. They fail `EXCESS_FRACTIONAL_PRECISION` and
`KEY_HAS_INVALID_CANDIDATE`; they must not be rounded into valid transactions or
classified as redundant. Every occurrence of a conflicting key must
be quarantined, including the original baseline record.

## Run the row-level examples

After installing SheetMerge as described in the project README, run from the
repository root:

```powershell
python -m sheetmerge --config samples/sources.toml --input samples/input/processable --output outputs/sample-review
```

Expected exit code: **1**, with a review output bundle. The directory must not
already exist. [expected.json](expected.json) is the machine-readable fixture
contract, also checked by pytest:

```text
33 input = 16 accepted + 14 rejected + 3 redundant
```

The 14 rejected records comprise six conflicting occurrences, three invalid dates,
three missing amounts, and two excess-precision candidates. Three valid exact
duplicates are reported as redundant, with
links to their accepted representatives.

| Currency | Input known amount | Accepted | Rejected known amount | Redundant |
| --- | ---: | ---: | ---: | ---: |
| USD | 356.55 | 141.45 | 115.00 | 100.10 |
| PLN | 583.35 | 220.35 | 163.00 | 200.00 |
| JPY | 7389 | 1290 | 5099 | 1000 |

Known amounts reconcile exactly in each currency. Each currency has one blank
amount; PLN also has two amounts that cannot be represented in minor units without
rounding. Five records are therefore unquantified and `monetary_complete` is false.
Unquantified amounts are never substituted with zero. The genuine zero-value transactions in
A and B are accepted.

## Run the structural failure example

```powershell
python -m sheetmerge --config samples/sources.toml --input samples/input --output outputs/sample-structural-failure
```

Expected exit code: **2**. The error names the missing-column workbook and both
missing headers; no completed output directory is published. The configuration
still requires all six canonical fields and does not fabricate missing values.

The malformed workbook is in a separate subdirectory so both workflows can be
tested without moving, editing, or removing any source file. The complete set has
36 physical data records, but the structurally invalid batch has no consolidation
result; the 33-record expectation applies only to `input/processable/`.

