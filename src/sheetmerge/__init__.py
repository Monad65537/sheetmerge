"""SheetMerge's ingestion, consolidation, reconciliation, and export API."""

from .config import Config, ConfigurationError, InputLimits, SourceProfile, load_config
from .consolidate import ClassifiedRow, ConsolidatedBatch, consolidate
from .ingest import IngestionError, discover_inputs, load_batch
from .models import BatchResult, Issue, RowResult, SourceRef
from .reconciliation import ReconciliationError, ReconciliationReport, reconcile
from .reporting import ExportError, ExportResult, export_batch

__all__ = [
    "BatchResult",
    "ClassifiedRow",
    "ConsolidatedBatch",
    "Config",
    "ConfigurationError",
    "ExportError",
    "ExportResult",
    "IngestionError",
    "InputLimits",
    "Issue",
    "ReconciliationError",
    "ReconciliationReport",
    "RowResult",
    "SourceProfile",
    "SourceRef",
    "consolidate",
    "discover_inputs",
    "export_batch",
    "load_batch",
    "load_config",
    "reconcile",
]
