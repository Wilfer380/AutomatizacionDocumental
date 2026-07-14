from __future__ import annotations

from pathlib import Path
import re
import shutil

from ..dossier_models import (
    DossierActionResult,
    DossierActionType,
    DossierExecutionMode,
    DossierStatus,
    DossierRunSummary,
)
from .dossier_backup_service import DossierBackupService
from ..utils.text import normalize_for_match


_NUMERIC_PREFIX_PATTERN = re.compile(r"^\s*\d+(?:\.\d+)+\s+")


class RealModeConfirmationRequiredError(RuntimeError):
    pass


class DossierDistributionService:
    def __init__(self, backup_service: DossierBackupService | None = None) -> None:
        self.backup_service = backup_service or DossierBackupService()

    def execute_plan(
        self,
        summary: DossierRunSummary,
        *,
        confirm_real: bool,
        replace_existing: bool = True,
        skip_if_destination_exists: bool = False,
        backup_root: Path | None = None,
        progress_callback=None,
    ) -> DossierRunSummary:
        if not confirm_real:
            raise RealModeConfirmationRequiredError("La distribución real requiere confirmación explícita antes de escribir archivos.")

        executed_items: list[DossierActionResult] = []
        total_actions = len(summary.items)
        for index, item in enumerate(summary.items, start=1):
            if item.status in {DossierStatus.ERROR, DossierStatus.SKIPPED} or not item.source_pdf_path or not item.planned_path:
                executed_items.append(item)
                if progress_callback:
                    progress_callback(index, total_actions, f"Acción {index} de {total_actions}")
                continue

            execution = self.copy_or_replace(
                Path(item.source_pdf_path),
                Path(item.planned_path),
                confirm_real=True,
                replace_existing=replace_existing,
                skip_if_destination_exists=True,
                backup_root=backup_root,
            )
            executed_items.append(self._merge_plan_and_execution(item, execution))
            if progress_callback:
                progress_callback(index, total_actions, f"Acción {index} de {total_actions}")

        summary.items = executed_items
        summary.execution_mode = DossierExecutionMode.REAL
        summary.errors = sum(1 for item in executed_items if item.status == DossierStatus.ERROR)
        summary.warnings = sum(1 for item in executed_items if item.status == DossierStatus.SKIPPED)
        summary.planned_actions = len(executed_items)
        return summary

    def copy_or_replace(
        self,
        source_pdf_path: Path,
        destination_path: Path,
        *,
        confirm_real: bool,
        replace_existing: bool = True,
        skip_if_destination_exists: bool = False,
        backup_root: Path | None = None,
    ) -> DossierActionResult:
        if not confirm_real:
            raise RealModeConfirmationRequiredError("La distribución real requiere confirmación explícita antes de escribir archivos.")

        if not source_pdf_path.is_file():
            return DossierActionResult(
                row_number=0,
                cp="",
                serie="",
                rule_name="distribution",
                target_folder=str(destination_path.parent),
                planned_path=str(destination_path),
                source_pdf_path=str(source_pdf_path),
                action_type=DossierActionType.SKIPPED,
                execution_mode=DossierExecutionMode.REAL,
                status=DossierStatus.ERROR,
                skipped_reason="El PDF de origen no existe.",
                observation="Falta el archivo de origen.",
            )

        destination_path.parent.mkdir(parents=True, exist_ok=True)
        backup_root = backup_root or destination_path.parent / "_backups"

        matched_final_path = destination_path if destination_path.exists() else self._find_equivalent_destination(destination_path)
        legacy_path = self._find_folder5_legacy_path(destination_path)
        final_document_exists = matched_final_path is not None

        if final_document_exists and skip_if_destination_exists:
            return DossierActionResult(
                row_number=0,
                cp="",
                serie="",
                rule_name="distribution",
                target_folder=str(destination_path.parent),
                planned_path=str(destination_path),
                source_pdf_path=str(source_pdf_path),
                action_type=DossierActionType.SKIPPED,
                execution_mode=DossierExecutionMode.REAL,
                status=DossierStatus.SKIPPED,
                skipped_reason=f"El documento final ya existe: {matched_final_path.name if matched_final_path else destination_path.name}",
                observation=f"No se puede agregar ni reemplazar porque el documento ya está correcto en {destination_path.parent}.",
            )

        if legacy_path is not None and not replace_existing:
            return DossierActionResult(
                row_number=0,
                cp="",
                serie="",
                rule_name="distribution",
                target_folder=str(destination_path.parent),
                planned_path=str(destination_path),
                source_pdf_path=str(source_pdf_path),
                action_type=DossierActionType.SKIPPED,
                execution_mode=DossierExecutionMode.REAL,
                status=DossierStatus.SKIPPED,
                skipped_reason=f"No se reemplazó porque ya existe el documento legacy: {legacy_path.name}",
                observation=f"No se puede reemplazar porque ya existe un documento en {destination_path.parent} y la opción de reemplazo está desactivada.",
            )

        existing_targets: list[Path] = []
        if matched_final_path is not None:
            existing_targets.append(matched_final_path)
        if legacy_path is not None and legacy_path not in existing_targets:
            existing_targets.append(legacy_path)

        backup_paths: list[str] = []
        for existing_path in existing_targets:
            backup_plan = self.backup_service.plan_destination_backup(existing_path, backup_root, simulated=False)
            if backup_plan.status == "copied":
                backup_paths.append(str(backup_plan.destination_path))

        action_type = DossierActionType.COPY
        status = DossierStatus.COPIED
        observation = f"Agregado correctamente en {destination_path.parent}."
        if existing_targets:
            action_type = DossierActionType.REPLACE
            status = DossierStatus.REPLACED
            if legacy_path is not None and legacy_path != destination_path:
                observation = f"Reemplazado correctamente {legacy_path.name} por {destination_path.name} en {destination_path.parent}."
            else:
                observation = f"Reemplazado correctamente en {destination_path.parent}."

        shutil.copy2(source_pdf_path, destination_path)
        if legacy_path is not None and legacy_path.exists() and legacy_path != destination_path:
            legacy_path.unlink()
        if matched_final_path is not None and matched_final_path.exists() and matched_final_path != destination_path:
            matched_final_path.unlink()

        return DossierActionResult(
            row_number=0,
            cp="",
            serie="",
            rule_name="distribution",
            target_folder=str(destination_path.parent),
            planned_path=str(destination_path),
            source_pdf_path=str(source_pdf_path),
            action_type=action_type,
            execution_mode=DossierExecutionMode.REAL,
            backup_path="; ".join(backup_paths),
            written_path=str(destination_path),
            status=status,
            observation=observation,
        )

    @staticmethod
    def _find_equivalent_destination(destination_path: Path) -> Path | None:
        if not destination_path.parent.exists():
            return None
        expected_key = normalize_for_match(destination_path.stem)
        expected_key_without_prefix = DossierDistributionService._normalized_without_numeric_prefix(destination_path.stem)
        if not expected_key and not expected_key_without_prefix:
            return None
        for candidate in destination_path.parent.glob("*.pdf"):
            if candidate == destination_path:
                continue
            candidate_key = normalize_for_match(candidate.stem)
            if expected_key and candidate_key == expected_key:
                return candidate
            candidate_key_without_prefix = DossierDistributionService._normalized_without_numeric_prefix(candidate.stem)
            if expected_key_without_prefix and candidate_key_without_prefix == expected_key_without_prefix:
                return candidate
        return None

    @staticmethod
    def _normalized_without_numeric_prefix(filename_stem: str) -> str:
        return normalize_for_match(_NUMERIC_PREFIX_PATTERN.sub("", filename_stem or "").strip())

    @staticmethod
    def _find_folder5_legacy_path(destination_path: Path) -> Path | None:
        file_key = normalize_for_match(destination_path.name)
        if "5 2 descriptivo de pintura pdf" not in file_key and "5 2 descriptivo de pintura" not in file_key:
            return None
        if not destination_path.parent.exists():
            return None
        for candidate in destination_path.parent.glob("*.pdf"):
            if candidate == destination_path:
                continue
            candidate_key = normalize_for_match(candidate.stem)
            if not candidate_key.startswith("5 2 procedimiento aplicacion pintura"):
                continue
            if "tanques" not in candidate_key:
                continue
            return candidate
        return None

    @staticmethod
    def _merge_plan_and_execution(plan: DossierActionResult, execution: DossierActionResult) -> DossierActionResult:
        return DossierActionResult(
            row_number=plan.row_number,
            cp=plan.cp,
            serie=plan.serie,
            rule_name=plan.rule_name,
            target_folder=plan.target_folder,
            planned_path=plan.planned_path,
            source_pdf_path=plan.source_pdf_path or execution.source_pdf_path,
            action_type=execution.action_type,
            execution_mode=execution.execution_mode,
            backup_path=execution.backup_path,
            written_path=execution.written_path,
            skipped_reason=execution.skipped_reason,
            status=execution.status,
            observation=execution.observation or plan.observation,
            timestamp=execution.timestamp,
        )
