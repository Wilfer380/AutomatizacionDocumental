from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path

from .config import (
    DEFAULT_DOSSIER_EXCEL_PATH,
    DEFAULT_DOSSIER_ROOT_PATH,
    DOSSIER_CP_SYNONYMS,
    DOSSIER_SERIE_SYNONYMS,
)


class DossierStatus(str, Enum):
    VALID = "valid"
    WARNING = "warning"
    ERROR = "error"
    BLOCKED = "blocked"
    PLANNED = "planned"
    SIMULATED = "simulated"
    SKIPPED = "skipped"
    COPIED = "copied"
    REPLACED = "replaced"


class DossierExecutionMode(str, Enum):
    SIMULATION = "simulation"
    REAL = "real"


class DossierActionType(str, Enum):
    PLANNED = "planned"
    COPY = "copy"
    REPLACE = "replace"
    SKIPPED = "skipped"


@dataclass(slots=True)
class DossierPdfSource:
    document_name: str
    source_pdf_path: Path
    target_folder: str
    final_name_pattern: str = "{serie}.pdf"
    action_type: DossierActionType = DossierActionType.PLANNED
    mandatory: bool = True

    def to_dict(self) -> dict:
        return {
            "document_name": self.document_name,
            "source_pdf_path": "" if str(self.source_pdf_path) in {"", "."} else str(self.source_pdf_path),
            "target_folder": self.target_folder,
            "final_name_pattern": self.final_name_pattern,
            "action_type": self.action_type.value,
            "mandatory": self.mandatory,
        }


@dataclass(slots=True)
class DossierRule:
    order: int
    name: str
    target_folder: str
    pdf_name_pattern: str = "{serie}.pdf"
    enabled: bool = True
    notes: str = ""

    @classmethod
    def from_mapping(cls, data: dict) -> "DossierRule":
        return cls(
            order=int(data.get("order", 0)),
            name=str(data.get("name", data.get("target_folder", "Rule"))).strip() or "Rule",
            target_folder=str(data.get("target_folder", "")).strip(),
            pdf_name_pattern=str(data.get("pdf_name_pattern", "{serie}.pdf")).strip() or "{serie}.pdf",
            enabled=bool(data.get("enabled", True)),
            notes=str(data.get("notes", "")).strip(),
        )

    def to_dict(self) -> dict:
        return {
            "order": self.order,
            "name": self.name,
            "target_folder": self.target_folder,
            "pdf_name_pattern": self.pdf_name_pattern,
            "enabled": self.enabled,
            "notes": self.notes,
        }


@dataclass(slots=True)
class DossierConfig:
    root_path: Path = DEFAULT_DOSSIER_ROOT_PATH
    excel_path: Path = DEFAULT_DOSSIER_EXCEL_PATH
    sheet_name: str = ""
    dossier_folder_name: str = "06_DOSSIER"
    simulation_only: bool = True
    replace_existing: bool = True
    cp_filter: str = ""
    serie_filter: str = ""
    cp_synonyms: tuple[str, ...] = DOSSIER_CP_SYNONYMS
    serie_synonyms: tuple[str, ...] = DOSSIER_SERIE_SYNONYMS
    pdf_sources: list[DossierPdfSource] = field(default_factory=list)
    rules: list[DossierRule] = field(default_factory=list)
    config_path: Path | None = None

    @classmethod
    def from_mapping(cls, data: dict, source_path: Path | None = None) -> "DossierConfig":
        pdf_sources = []
        for item in data.get("pdf_sources", []):
            if isinstance(item, dict):
                source_pdf_path = str(item.get("source_pdf_path", "")).strip()
                action_type_raw = str(item.get("action_type", DossierActionType.PLANNED.value)).strip().lower()
                action_type_aliases = {
                    "agregar": DossierActionType.COPY.value,
                    "copiar": DossierActionType.COPY.value,
                    "replace": DossierActionType.REPLACE.value,
                    "reemplazar": DossierActionType.REPLACE.value,
                    "planificado": DossierActionType.PLANNED.value,
                    "omitido": DossierActionType.SKIPPED.value,
                }
                pdf_sources.append(
                    DossierPdfSource(
                        document_name=str(item.get("document_name", "")).strip(),
                        source_pdf_path=Path(source_pdf_path).expanduser() if source_pdf_path else Path(),
                        target_folder=str(item.get("target_folder", "")).strip(),
                        final_name_pattern=str(item.get("final_name_pattern", "{serie}.pdf")).strip() or "{serie}.pdf",
                        action_type=DossierActionType(action_type_aliases.get(action_type_raw, action_type_raw) or DossierActionType.PLANNED.value),
                        mandatory=bool(item.get("mandatory", True)),
                    )
                )
        rules = [DossierRule.from_mapping(item) for item in data.get("rules", []) if isinstance(item, dict)]
        if not rules and pdf_sources:
            rules = [
                DossierRule(
                    order=index,
                    name=source.document_name or f"Documento {index}",
                    target_folder=source.target_folder,
                    pdf_name_pattern=source.final_name_pattern,
                    enabled=True,
                    notes=f"source={source.source_pdf_path}; mandatory={source.mandatory}",
                )
                for index, source in enumerate(pdf_sources, start=1)
                if source.target_folder
            ]
        return cls(
            root_path=Path(data.get("root_path", DEFAULT_DOSSIER_ROOT_PATH)).expanduser(),
            excel_path=Path(data.get("excel_path", DEFAULT_DOSSIER_EXCEL_PATH)).expanduser(),
            sheet_name=str(data.get("sheet_name", "")).strip(),
            dossier_folder_name=str(data.get("dossier_folder_name", "06_DOSSIER")).strip() or "06_DOSSIER",
            simulation_only=bool(data.get("simulation_only", True)),
            replace_existing=bool(data.get("replace_existing", True)),
            cp_filter=str(data.get("cp_filter", "")).strip(),
            serie_filter=str(data.get("serie_filter", "")).strip(),
            cp_synonyms=tuple(str(item).strip() for item in data.get("cp_synonyms", DOSSIER_CP_SYNONYMS) if str(item).strip()),
            serie_synonyms=tuple(str(item).strip() for item in data.get("serie_synonyms", DOSSIER_SERIE_SYNONYMS) if str(item).strip()),
            pdf_sources=pdf_sources,
            rules=sorted(rules, key=lambda rule: (rule.order, rule.name.lower())),
            config_path=source_path,
        )

    def to_dict(self) -> dict:
        return {
            "root_path": str(self.root_path),
            "excel_path": str(self.excel_path),
            "sheet_name": self.sheet_name,
            "dossier_folder_name": self.dossier_folder_name,
            "simulation_only": self.simulation_only,
            "replace_existing": self.replace_existing,
            "cp_filter": self.cp_filter,
            "serie_filter": self.serie_filter,
            "cp_synonyms": list(self.cp_synonyms),
            "serie_synonyms": list(self.serie_synonyms),
            "pdf_sources": [source.to_dict() for source in self.pdf_sources],
            "rules": [rule.to_dict() for rule in self.rules],
            "config_path": str(self.config_path) if self.config_path else "",
        }


@dataclass(slots=True)
class DossierColumnInfo:
    index: int
    name: str


@dataclass(slots=True)
class DossierWorkbookInfo:
    path: Path
    sheet_name: str
    header_row: int
    columns: list[DossierColumnInfo] = field(default_factory=list)
    cp_column: str = ""
    serie_column: str = ""


@dataclass(slots=True)
class DossierRow:
    row_number: int
    cp: str
    serie: str
    normalized_cp: str
    normalized_serie: str
    status: DossierStatus = DossierStatus.VALID
    observation: str = ""
    cp_folder: str = ""
    dossier_folder: str = ""
    planos_folder: str = ""
    folder_5: str = ""
    folder_6: str = ""
    folder_7: str = ""
    series_in_planos: bool = False
    matched_dossier_folders: list[str] = field(default_factory=list)
    matched_planos_folders: list[str] = field(default_factory=list)
    matched_folder_5: list[str] = field(default_factory=list)
    matched_folder_6: list[str] = field(default_factory=list)
    matched_folder_7: list[str] = field(default_factory=list)


@dataclass(slots=True)
class DossierActionResult:
    row_number: int
    cp: str
    serie: str
    rule_name: str
    target_folder: str
    planned_path: str
    source_pdf_path: str = ""
    action_type: DossierActionType = DossierActionType.PLANNED
    execution_mode: DossierExecutionMode = DossierExecutionMode.SIMULATION
    backup_path: str = ""
    simulation_path: str = ""
    written_path: str = ""
    skipped_reason: str = ""
    status: DossierStatus = DossierStatus.PLANNED
    observation: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat(timespec="seconds"))

    def to_dict(self) -> dict:
        return {
            "row_number": self.row_number,
            "cp": self.cp,
            "serie": self.serie,
            "rule_name": self.rule_name,
            "target_folder": self.target_folder,
            "planned_path": self.planned_path,
            "source_pdf_path": self.source_pdf_path,
            "action_type": self.action_type.value,
            "execution_mode": self.execution_mode.value,
            "backup_path": self.backup_path,
            "simulation_path": self.simulation_path,
            "written_path": self.written_path,
            "skipped_reason": self.skipped_reason,
            "status": self.status.value,
            "observation": self.observation,
            "timestamp": self.timestamp,
        }


@dataclass(slots=True)
class DossierRunSummary:
    total_rows: int
    valid_rows: int
    warnings: int
    errors: int
    blocked: int = 0
    planned_actions: int = 0
    execution_mode: DossierExecutionMode = DossierExecutionMode.SIMULATION
    report_path: Path | None = None
    items: list[DossierActionResult] = field(default_factory=list)
    backup_path: str = ""
    simulation_root: str = ""

    def to_dict(self) -> dict:
        return {
            "total_rows": self.total_rows,
            "valid_rows": self.valid_rows,
            "warnings": self.warnings,
            "errors": self.errors,
            "blocked": self.blocked,
            "planned_actions": self.planned_actions,
            "execution_mode": self.execution_mode.value,
            "report_path": str(self.report_path) if self.report_path else "",
            "backup_path": self.backup_path,
            "simulation_root": self.simulation_root,
            "items": [item.to_dict() for item in self.items],
        }
