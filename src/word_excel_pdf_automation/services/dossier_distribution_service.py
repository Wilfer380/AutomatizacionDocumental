from __future__ import annotations

from pathlib import Path
import shutil

from ..dossier_models import (
    DossierActionResult,
    DossierActionType,
    DossierExecutionMode,
    DossierStatus,
    DossierRunSummary,
)
from .dossier_backup_service import DossierBackupService


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
        backup_root: Path | None = None,
        progress_callback=None,
    ) -> DossierRunSummary:
        if not confirm_real:
            raise RealModeConfirmationRequiredError("La distribuci?n real requiere confirmaci?n expl?cita antes de escribir archivos.")

        executed_items: list[DossierActionResult] = []
        total_actions = len(summary.items)
        for index, item in enumerate(summary.items, start=1):
            if item.status in {DossierStatus.ERROR, DossierStatus.SKIPPED} or not item.source_pdf_path or not item.planned_path:
                executed_items.append(item)
                if progress_callback:
                    progress_callback(index, total_actions, f"Acci?n {index} de {total_actions}")
                continue

            execution = self.copy_or_replace(
                Path(item.source_pdf_path),
                Path(item.planned_path),
                confirm_real=True,
                replace_existing=replace_existing,
                backup_root=backup_root,
            )
            executed_items.append(self._merge_plan_and_execution(item, execution))
            if progress_callback:
                progress_callback(index, total_actions, f"Acci?n {index} de {total_actions}")

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
        backup_root: Path | None = None,
    ) -> DossierActionResult:
        if not confirm_real:
            raise RealModeConfirmationRequiredError("La distribuci?n real requiere confirmaci?n expl?cita antes de escribir archivos.")

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

        legacy_path = self._find_folder5_legacy_path(destination_path)
        existing_targets: list[Path] = []
        if destination_path.exists():
            existing_targets.append(destination_path)
        if legacy_path is not None and legacy_path not in existing_targets:
            existing_targets.append(legacy_path)

        if existing_targets and not replace_existing:
            existing_names = ", ".join(path.name for path in existing_targets)
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
                skipped_reason=f"No se agreg? porque ya existe: {existing_names}",
                observation=f"No se puede agregar ya que los documentos est?n agregados en {destination_path.parent}.",
            )

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
    def _find_folder5_legacy_path(destination_path: Path) -> Path | None:
        file_key = destination_path.name.lower()
        if "5.2 descriptivo de pintura" not in file_key:
            return None
        if not destination_path.parent.exists():
            return None
        for candidate in destination_path.parent.glob("*.pdf"):
            if candidate == destination_path:
                continue
            candidate_key = candidate.stem.lower()
            normalized = "".join(ch if ch.isalnum() else " " for ch in candidate_key)
            normalized = " ".join(normalized.split())
            if normalized.startswith("5 2 procedimiento aplicacion pintura"):
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
