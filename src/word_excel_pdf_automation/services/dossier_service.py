from __future__ import annotations

import logging
from pathlib import Path

from ..config import DEFAULT_DOSSIER_CONFIG_EXAMPLE_PATH, DEFAULT_DOSSIER_CONFIG_PATH, DEFAULT_DOSSIER_SIMULATION_DIR, DEFAULT_OUTPUT_DIR, read_json_file
from ..dossier_models import DossierActionResult, DossierActionType, DossierConfig, DossierExecutionMode, DossierRunSummary, DossierStatus
from ..utils.text import normalize_for_match
from .dossier_backup_service import DossierBackupService
from .dossier_distribution_service import DossierDistributionService, RealModeConfirmationRequiredError
from .dossier_file_placer_service import DossierFilePlacerService
from .dossier_report_service import DossierReportService
from .dossier_simulation_workspace_service import DossierSimulationWorkspaceService
from .dossier_validator_service import DossierValidatorService, DossierWorkbookHeaderError


logger = logging.getLogger(__name__)


class DossierService:
    def __init__(self, validator: DossierValidatorService | None = None, placer: DossierFilePlacerService | None = None, report_service: DossierReportService | None = None, backup_service: DossierBackupService | None = None, distribution_service: DossierDistributionService | None = None, simulation_workspace_service: DossierSimulationWorkspaceService | None = None) -> None:
        self.validator = validator or DossierValidatorService()
        self.placer = placer or DossierFilePlacerService()
        self.report_service = report_service or DossierReportService()
        self.backup_service = backup_service or DossierBackupService()
        self.distribution_service = distribution_service or DossierDistributionService(self.backup_service)
        self.simulation_workspace_service = simulation_workspace_service or DossierSimulationWorkspaceService()

    def load_config(self, config_path: Path | None = None) -> DossierConfig:
        resolved_path = config_path or DEFAULT_DOSSIER_CONFIG_PATH
        payload = read_json_file(resolved_path) or read_json_file(DEFAULT_DOSSIER_CONFIG_EXAMPLE_PATH)
        return DossierConfig.from_mapping(payload, source_path=resolved_path)

    def run_simulation(self, config: DossierConfig, report_dir: Path | None = None, progress_callback=None) -> DossierRunSummary:
        summary = self._build_planned_summary(config, progress_callback=progress_callback)
        summary.execution_mode = DossierExecutionMode.SIMULATION
        if not summary.blocked:
            try:
                backup = self.backup_service.plan_backup(config.config_path, DEFAULT_OUTPUT_DIR / 'Backups', simulated=True)
                summary.backup_path = str(backup.destination_path)
            except Exception as exc:
                logger.exception("No se pudo preparar el respaldo de simulación")
            try:
                summary = self.simulation_workspace_service.materialize(config.root_path, summary, DEFAULT_DOSSIER_SIMULATION_DIR)
            except Exception as exc:
                logger.exception("No se pudo materializar la simulación")
        report_root = report_dir or DEFAULT_OUTPUT_DIR / 'DossierReports'
        try:
            self.report_service.write_json_report(report_root, summary)
        except Exception as exc:
            logger.exception("No se pudo escribir el reporte de simulación")
        return summary

    def run_distribution(self, config: DossierConfig, confirm_real: bool = False, report_dir: Path | None = None, phase1_items: list | None = None, progress_callback=None) -> DossierRunSummary:
        summary = self._build_planned_summary(config, progress_callback=progress_callback)
        if summary.blocked:
            summary.execution_mode = DossierExecutionMode.REAL if not config.simulation_only else DossierExecutionMode.SIMULATION
            report_root = report_dir or DEFAULT_OUTPUT_DIR / 'DossierReports'
            self.report_service.write_json_report(report_root, summary)
            return summary
        if config.simulation_only:
            summary.execution_mode = DossierExecutionMode.SIMULATION
        else:
            if not confirm_real:
                raise RealModeConfirmationRequiredError('La distribución real requiere confirmación explícita antes de escribir archivos.')
            summary = self.distribution_service.execute_plan(summary, confirm_real=True, replace_existing=config.replace_existing, progress_callback=progress_callback)
        if phase1_items:
            phase1_routes = self._build_phase1_routes(config, summary, phase1_items)
            summary.items.extend(phase1_routes)
            summary.errors = sum(1 for item in summary.items if item.status == DossierStatus.ERROR)
            summary.warnings = sum(1 for item in summary.items if item.status == DossierStatus.SKIPPED)
            summary.planned_actions = len(summary.items)
        if summary.execution_mode == DossierExecutionMode.SIMULATION:
            try:
                summary = self.simulation_workspace_service.materialize(config.root_path, summary, DEFAULT_DOSSIER_SIMULATION_DIR)
            except Exception as exc:
                logger.exception("No se pudo materializar la simulación")
        try:
            backup = self.backup_service.plan_backup(config.config_path, DEFAULT_OUTPUT_DIR / 'Backups', simulated=config.simulation_only)
            summary.backup_path = str(backup.destination_path)
        except Exception:
            logger.exception("No se pudo preparar el respaldo de la ejecución")
        report_root = report_dir or DEFAULT_OUTPUT_DIR / 'DossierReports'
        try:
            self.report_service.write_json_report(report_root, summary)
        except Exception:
            logger.exception("No se pudo escribir el reporte de dossier")
        return summary

    def _build_planned_summary(self, config: DossierConfig, progress_callback=None) -> DossierRunSummary:
        try:
            _workbook_info, rows = self.validator.load_rows(config)
        except DossierWorkbookHeaderError as exc:
            blocked_item = DossierActionResult(row_number=0, cp='', serie='', rule_name='validation', target_folder='', planned_path='', action_type=DossierActionType.SKIPPED, status=DossierStatus.BLOCKED, skipped_reason=str(exc), observation=f'Bloqueado: {exc}')
            return DossierRunSummary(total_rows=0, valid_rows=0, warnings=0, errors=0, blocked=1, planned_actions=1, items=[blocked_item])
        rows = self._filter_rows(config, rows)
        validated_rows = self.validator.validate_paths(config, rows, progress_callback=progress_callback)
        return self.placer.build_simulation(config, validated_rows, progress_callback=progress_callback)

    def _filter_rows(self, config: DossierConfig, rows: list) -> list:
        cp_filter = normalize_for_match(config.cp_filter)
        serie_filter = normalize_for_match(config.serie_filter)
        if not cp_filter and not serie_filter:
            return rows

        filtered_rows = []
        for row in rows:
            cp_matches = not cp_filter or cp_filter in (row.normalized_cp or '')
            serie_matches = not serie_filter or serie_filter in (row.normalized_serie or '')
            if cp_matches and serie_matches:
                filtered_rows.append(row)
        return filtered_rows

    def _build_phase1_routes(self, config: DossierConfig, summary: DossierRunSummary, phase1_items: list) -> list[DossierActionResult]:
        from ..utils.text import normalize_for_match

        dossier_targets_by_row: dict[int, list[tuple[str, str, str]]] = {}
        seen_targets: set[tuple[int, str]] = set()
        for item in summary.items:
            if item.row_number <= 0 or not item.target_folder:
                continue
            dossier_folder = str(Path(item.target_folder).parent)
            if not dossier_folder:
                continue
            dedupe_key = (item.row_number, dossier_folder)
            if dedupe_key in seen_targets:
                continue
            seen_targets.add(dedupe_key)
            dossier_targets_by_row.setdefault(item.row_number, []).append((item.cp, item.serie, dossier_folder))

        routed_items: list[DossierActionResult] = []
        for generation in phase1_items:
            candidate_targets = dossier_targets_by_row.get(getattr(generation, 'row_number', 0), [])
            pdf_path = Path(getattr(generation, 'pdf_path', '') or getattr(generation, 'pdf_filename', ''))
            pdf_filename = getattr(generation, 'pdf_filename', '') or pdf_path.name
            generation_series = getattr(generation, 'series', '')
            if not candidate_targets:
                routed_items.append(DossierActionResult(row_number=getattr(generation, 'row_number', 0), cp='', serie=generation_series, rule_name='phase1-routing', target_folder='', planned_path='', source_pdf_path=str(pdf_path), action_type=DossierActionType.SKIPPED, execution_mode=DossierExecutionMode.SIMULATION, status=DossierStatus.SKIPPED, skipped_reason='No se encontr? una fila validada del dossier para la ruta.', observation='Ruta de Fase 1 omitida.'))
                continue
            for cp_value, serie_value, dossier_folder_value in candidate_targets:
                if normalize_for_match(serie_value) != normalize_for_match(generation_series):
                    routed_items.append(DossierActionResult(row_number=getattr(generation, 'row_number', 0), cp=cp_value, serie=serie_value, rule_name='phase1-routing', target_folder='', planned_path='', source_pdf_path=str(pdf_path), action_type=DossierActionType.SKIPPED, execution_mode=DossierExecutionMode.SIMULATION, status=DossierStatus.SKIPPED, skipped_reason='La serie del archivo de Fase 1 no coincide con la fila del dossier.', observation='Ruta de Fase 1 omitida.'))
                    continue
                destination_folder = self.validator.find_folder_6(Path(dossier_folder_value))
                if not destination_folder:
                    routed_items.append(DossierActionResult(row_number=getattr(generation, 'row_number', 0), cp=cp_value, serie=serie_value, rule_name='phase1-routing', target_folder='', planned_path='', source_pdf_path=str(pdf_path), action_type=DossierActionType.SKIPPED, execution_mode=DossierExecutionMode.SIMULATION, status=DossierStatus.SKIPPED, skipped_reason='No fue posible resolver la carpeta 6 del dossier.', observation='Ruta de Fase 1 omitida.'))
                    continue
                planned_path = destination_folder / pdf_filename
                if not pdf_path.is_file():
                    routed_items.append(DossierActionResult(row_number=getattr(generation, 'row_number', 0), cp=cp_value, serie=serie_value, rule_name='phase1-routing', target_folder=str(destination_folder), planned_path=str(planned_path), source_pdf_path=str(pdf_path), action_type=DossierActionType.SKIPPED, execution_mode=DossierExecutionMode.SIMULATION, status=DossierStatus.WARNING, skipped_reason='PDF generado por Fase 1 no encontrado para la serie.', observation='Ruta de Fase 1 omitida.'))
                    continue
                planned_route = DossierActionResult(row_number=getattr(generation, 'row_number', 0), cp=cp_value, serie=serie_value, rule_name='phase1-routing', target_folder=str(destination_folder), planned_path=str(planned_path), source_pdf_path=str(pdf_path), action_type=DossierActionType.PLANNED, execution_mode=DossierExecutionMode.SIMULATION, status=DossierStatus.PLANNED, observation='Ruta de Fase 1 planificada hacia 6 Trazabilidad.')
                if config.simulation_only:
                    routed_items.append(planned_route)
                else:
                    routed_items.append(
                        self.distribution_service._merge_plan_and_execution(
                            planned_route,
                            self.distribution_service.copy_or_replace(
                                pdf_path,
                                planned_path,
                                confirm_real=True,
                                replace_existing=config.replace_existing,
                            ),
                        )
                    )
        return routed_items
