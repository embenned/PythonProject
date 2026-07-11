from dataclasses import dataclass, field
from typing import Any


class ComparisonError(Exception):
    """Base exception for workbook comparison failures."""


class WorkbookLoadError(ComparisonError):
    """Raised when one or both workbooks cannot be loaded."""


class ReportGenerationError(ComparisonError):
    """Raised when the deviation report cannot be written."""


@dataclass(slots=True)
class DeviationRecord:
    change_type: str
    sheet_name: str
    cell_address: str
    standard_value: Any
    project_value: Any


@dataclass(slots=True)
class ComparisonSummary:
    total_sheets_compared: int = 0
    total_cells_compared: int = 0
    total_deviations_found: int = 0
    execution_time_seconds: float = 0.0


@dataclass(slots=True)
class WorkbookComparisonResult:
    deviations: list[DeviationRecord] = field(default_factory=list)
    summary: ComparisonSummary = field(default_factory=ComparisonSummary)
    output_file: str = ""
