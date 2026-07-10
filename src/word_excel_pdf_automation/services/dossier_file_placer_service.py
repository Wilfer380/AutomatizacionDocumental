from __future__ import annotations

from pathlib import Path

from ..dossier_models import DossierActionResult, DossierActionType, DossierConfig, DossierPdfSource, DossierRow, DossierRunSummary, DossierStatus
from ..utils.text import normalize_for_match
from .dossier_sequence_service import DossierSequenceService


class DossierFilePlacerService:
    def __init__(self, sequence_service: DossierSequenceService | None = None) -> None:
        self.sequence_service = sequence_service or DossierSequenceService()

    def build_simulation(self, config: DossierConfig, rows: list[DossierRow], progress_callback=None) -> DossierRunSummary:
        items: list[DossierActionResult] = []
        warnings = 0
        errors = 0
        blocked = 0
        valid_rows = 0
        source_lookup = self._build_source_lookup(config.pdf_sources)
        total_actions = self._estimate_planned_actions(config, rows)
        for row in rows:
            if row.status == DossierStatus.BLOCKED:
                blocked += 1
                items.append(self._build_validation_result(row, DossierStatus.BLOCKED))
                self._notify(progress_callback, len(items), total_actions, row)
                continue
            if row.status == DossierStatus.ERROR:
                errors += 1
                items.append(self._build_validation_result(row, DossierStatus.ERROR))
                self._notify(progress_callback, len(items), total_actions, row)
                continue
            if row.status == DossierStatus.SKIPPED or not row.matched_dossier_folders:
                warnings += 1
                items.append(self._build_validation_result(row, DossierStatus.SKIPPED))
                self._notify(progress_callback, len(items), total_actions, row)
                continue
            valid_rows += 1
            for target in self._iter_targets(row):
                for rule in self.sequence_service.sort_rules(config.rules):
                    if not self._is_phase2_folder(rule.target_folder):
                        continue
                    actual_target_folder = self._resolve_target_folder(rule.target_folder, target)
                    if actual_target_folder is None:
                        warnings += 1
                        items.append(DossierActionResult(row_number=row.row_number, cp=row.cp, serie=row.serie, rule_name=rule.name, target_folder="", planned_path="", action_type=DossierActionType.SKIPPED, status=DossierStatus.SKIPPED, skipped_reason="No se encontr? la carpeta destino dentro del dossier.", observation="Falta la carpeta destino durante la planificaci?n."))
                        self._notify(progress_callback, len(items), total_actions, row)
                        continue
                    source_pdf = source_lookup.get(normalize_for_match(rule.name)) or source_lookup.get(normalize_for_match(rule.target_folder))
                    source_path = Path(source_pdf.source_pdf_path) if source_pdf else Path()
                    if not source_pdf or not source_path.is_file():
                        missing_status = DossierStatus.ERROR if (source_pdf is None or source_pdf.mandatory) else DossierStatus.SKIPPED
                        if missing_status == DossierStatus.ERROR:
                            errors += 1
                        else:
                            warnings += 1
                        items.append(DossierActionResult(row_number=row.row_number, cp=row.cp, serie=row.serie, rule_name=rule.name, target_folder=str(actual_target_folder), planned_path=str(actual_target_folder / self.sequence_service.build_destination_filename(row, rule)), source_pdf_path=str(source_path) if source_pdf else "", action_type=DossierActionType.SKIPPED, status=missing_status, skipped_reason="Falta el PDF de origen obligatorio." if missing_status == DossierStatus.ERROR else "Falta el PDF de origen opcional.", observation="Falta el PDF de origen durante la planificaci?n."))
                        self._notify(progress_callback, len(items), total_actions, row)
                        continue
                    effective_rule = rule
                    if source_pdf.final_name_pattern:
                        effective_rule = type(rule)(order=rule.order, name=rule.name, target_folder=rule.target_folder, pdf_name_pattern=source_pdf.final_name_pattern, enabled=rule.enabled, notes=rule.notes)
                    planned_path = actual_target_folder / self.sequence_service.build_destination_filename(row, effective_rule)
                    existing_path, replace_note = self._detect_existing_destination(rule, planned_path)
                    if existing_path and not config.replace_existing:
                        warnings += 1
                        items.append(DossierActionResult(row_number=row.row_number, cp=row.cp, serie=row.serie, rule_name=rule.name, target_folder=str(actual_target_folder), planned_path=str(planned_path), source_pdf_path=str(source_pdf.source_pdf_path), action_type=DossierActionType.SKIPPED, status=DossierStatus.SKIPPED, skipped_reason=f"Ya existe un documento destino: {existing_path.name}", observation=f"No se puede agregar ya que el documento ya est? agregado en {existing_path.parent}."))
                        self._notify(progress_callback, len(items), total_actions, row)
                        continue
                    observation = "Planificado en modo simulaci?n"
                    action_type = DossierActionType.PLANNED
                    if existing_path:
                        observation = replace_note or f"Se reemplazar? el documento existente {existing_path.name}."
                        action_type = DossierActionType.REPLACE
                    items.append(DossierActionResult(row_number=row.row_number, cp=row.cp, serie=row.serie, rule_name=rule.name, target_folder=str(actual_target_folder), planned_path=str(planned_path), source_pdf_path=str(source_pdf.source_pdf_path), action_type=action_type, status=DossierStatus.PLANNED, observation=observation))
                    self._notify(progress_callback, len(items), total_actions, row)
        return DossierRunSummary(total_rows=len(rows), valid_rows=valid_rows, warnings=warnings, errors=errors, blocked=blocked, planned_actions=len(items), items=items)

    @staticmethod
    def _notify(progress_callback, current: int, total: int, row: DossierRow) -> None:
        if progress_callback:
            progress_callback(current, total, f"Fila {row.row_number} planificada")

    @staticmethod
    def _is_phase2_folder(target_folder: str) -> bool:
        normalized = normalize_for_match(target_folder)
        return normalized.startswith("5 ") or normalized == "5" or normalized.startswith("6 ") or normalized == "6" or normalized.startswith("7 ") or normalized == "7"

    @staticmethod
    def _build_source_lookup(sources: list[DossierPdfSource]) -> dict[str, DossierPdfSource]:
        lookup: dict[str, DossierPdfSource] = {}
        for source in sources:
            if source.target_folder:
                lookup.setdefault(normalize_for_match(source.target_folder), source)
            if source.document_name:
                lookup.setdefault(normalize_for_match(source.document_name), source)
        return lookup

    def _estimate_planned_actions(self, config: DossierConfig, rows: list[DossierRow]) -> int:
        total = 0
        for row in rows:
            if row.status in {DossierStatus.BLOCKED, DossierStatus.ERROR, DossierStatus.SKIPPED} or not row.matched_dossier_folders:
                total += 1
                continue
            total += len(self._iter_targets(row)) * sum(1 for rule in self.sequence_service.sort_rules(config.rules) if self._is_phase2_folder(rule.target_folder))
        return total

    @staticmethod
    def _build_validation_result(row: DossierRow, status: DossierStatus) -> DossierActionResult:
        return DossierActionResult(row_number=row.row_number, cp=row.cp, serie=row.serie, rule_name="validation", target_folder="", planned_path="", status=status, observation=row.observation)

    @staticmethod
    def _iter_targets(row: DossierRow) -> list[dict[str, str]]:
        size = max(len(row.matched_dossier_folders), len(row.matched_folder_5), len(row.matched_folder_6), len(row.matched_folder_7))
        return [{"dossier_folder": row.matched_dossier_folders[i] if i < len(row.matched_dossier_folders) else "", "folder_5": row.matched_folder_5[i] if i < len(row.matched_folder_5) else "", "folder_6": row.matched_folder_6[i] if i < len(row.matched_folder_6) else "", "folder_7": row.matched_folder_7[i] if i < len(row.matched_folder_7) else ""} for i in range(size)]

    @staticmethod
    def _detect_existing_destination(rule, planned_path: Path) -> tuple[Path | None, str]:
        if planned_path.exists():
            return planned_path, f"Se reemplazar? el documento existente {planned_path.name}."
        legacy_path = DossierFilePlacerService._find_folder5_legacy_path(rule, planned_path)
        if legacy_path is not None:
            return legacy_path, f"Se reemplazar? {legacy_path.name} por {planned_path.name}."
        return None, ""

    @staticmethod
    def _find_folder5_legacy_path(rule, planned_path: Path) -> Path | None:
        target_key = normalize_for_match(getattr(rule, 'target_folder', ''))
        rule_key = normalize_for_match(getattr(rule, 'name', ''))
        file_key = normalize_for_match(planned_path.name)
        if not (target_key.startswith('5') and 'descriptivo de pintura' in file_key and 'descriptivo de pintura' in rule_key):
            return None
        if not planned_path.parent.exists():
            return None
        for candidate in planned_path.parent.glob('*.pdf'):
            if candidate == planned_path:
                continue
            candidate_key = normalize_for_match(candidate.stem)
            if candidate_key.startswith('5 2 procedimiento aplicacion pintura'):
                return candidate
        return None

    @staticmethod
    def _resolve_target_folder(rule_target_folder: str, target: dict[str, str]) -> Path | None:
        normalized = normalize_for_match(rule_target_folder)
        if normalized.startswith("5"):
            return Path(target["folder_5"]) if target.get("folder_5") else None
        if normalized.startswith("6"):
            return Path(target["folder_6"]) if target.get("folder_6") else None
        if normalized.startswith("7"):
            return Path(target["folder_7"]) if target.get("folder_7") else None
        return Path(target["dossier_folder"]) if target.get("dossier_folder") else None
