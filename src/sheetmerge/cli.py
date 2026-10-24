"""Batch command with explicit review/failure exit statuses."""

import argparse
import json
import sys
from collections.abc import Sequence
from pathlib import Path

from .config import ConfigurationError, load_config
from .consolidate import consolidate
from .ingest import IngestionError, load_batch
from .reconciliation import ReconciliationError
from .reporting import ExportError, export_batch, validate_output_location
from .version import VERSION


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate and consolidate sales CSV/XLSX inputs")
    parser.add_argument("--version", action="version", version=f"SheetMerge {VERSION}")
    parser.add_argument("--config", type=Path, required=True, help="Explicit TOML source profiles")
    parser.add_argument("--input", type=Path, required=True, help="Dedicated input directory")
    parser.add_argument(
        "--output", type=Path, required=True, help="New directory outside input tree"
    )
    args = parser.parse_args(argv)
    try:
        config = load_config(args.config)
        validate_output_location(args.input, args.output)
        batch = load_batch(args.input, config)
        result = consolidate(batch)
        exported = export_batch(result, config, args.output)
    except (ConfigurationError, IngestionError, ReconciliationError, ExportError, OSError) as exc:
        print(f"sheetmerge: error: {exc}", file=sys.stderr)
        return 2
    report = exported.reconciliation
    review = bool(result.rejected) or not report.complete
    print(
        json.dumps(
            {
                "status": "review_required" if review else "completed",
                "output_directory": str(exported.output_directory),
                "counts": report.to_dict()["counts"],
                "count_balanced": True,
                "known_amounts_balanced": True,
                "monetary_complete": report.complete,
                "unquantified_count": report.unquantified_count,
            },
            sort_keys=True,
        )
    )
    return 1 if review else 0
