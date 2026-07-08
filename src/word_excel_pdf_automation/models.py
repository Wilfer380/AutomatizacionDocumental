from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Optional


class ValidationStatus(str, Enum):
    VALID = "Valid"
    EMPTY = "Empty"
    DUPLICATE = "Duplicate"
    INVALID = "Invalid"


class ConflictStrategy(str, Enum):
    OVERWRITE = "overwrite"
    SKIP = "skip"
    COUNTER = "counter"


@dataclass(slots=True)
class ColumnInfo:
    index: int
    name: str


@dataclass(slots=True)
class SheetInfo:
    name: str
    header_row: int
    columns: list[ColumnInfo]
    suggested_serial_column: Optional[str] = None


@dataclass(slots=True)
class WorkbookInfo:
    path: Path
    sheets: list[SheetInfo] = field(default_factory=list)


@dataclass(slots=True)
class SeriesRow:
    row_number: int
    series: str
    normalized_series: str
    selected: bool = True
    status: ValidationStatus = ValidationStatus.VALID
    observation: str = ""
    source_sheet: str = ""

    @property
    def is_empty(self) -> bool:
        return self.status == ValidationStatus.EMPTY

    @property
    def is_duplicate(self) -> bool:
        return self.status == ValidationStatus.DUPLICATE


@dataclass(slots=True)
class BatchOptions:
    template_path: Path
    excel_path: Path
    output_dir: Path
    sheet_name: str
    serial_column: str
    exclude_empties: bool = True
    exclude_duplicates: bool = True
    search_term: str = ""
    process_only_selected: bool = True
    libreoffice_path: str = ""
    placeholder: str = "[Serie]"
    conflict_strategy: ConflictStrategy = ConflictStrategy.OVERWRITE


@dataclass(slots=True)
class GenerationResult:
    row_number: int
    series: str
    docx_filename: str = ""
    pdf_filename: str = ""
    docx_path: str = ""
    pdf_path: str = ""
    status: str = ""
    observation: str = ""
    timestamp: str = ""


@dataclass(slots=True)
class BatchSummary:
    total: int
    succeeded: int
    failed: int
    skipped: int
    report_path: Path
    generated_items: list[GenerationResult] = field(default_factory=list)


@dataclass(slots=True)
class BatchProgress:
    current: int
    total: int
    message: str
