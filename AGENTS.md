Project: SheetMerge

Purpose:
Consolidate heterogeneous sales CSV/XLSX files into an
auditable, validated master dataset.

Rules:
- Use Python 3.12+, pandas, openpyxl, pytest, and Ruff.
- Never modify input files.
- Preserve identifier strings and leading zeroes.
- Avoid lossy monetary arithmetic.
- No silent row deletion or deduplication.
- Document assumptions and ambiguous cases.
- Use deterministic output ordering.
- Avoid unnecessary dependencies.
- Keep raw sample files synthetic.
- Add tests for behavior and edge cases.
- Do not run destructive Git commands.
- Do not commit or push without user approval.

Verification:
python -m pytest -q
python -m ruff check .