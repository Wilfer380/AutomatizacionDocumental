from __future__ import annotations

import logging
import os
import shutil
from datetime import datetime
from pathlib import Path

from ..dossier_models import DossierExecutionMode, DossierRunSummary, DossierStatus


logger = logging.getLogger(__name__)


class DossierSimulationWorkspaceService:
    def materialize(self, root_path: Path, summary: DossierRunSummary, simulation_root_base: Path) -> DossierRunSummary:
        if summary.execution_mode != DossierExecutionMode.SIMULATION:
            return summary

        simulation_root = simulation_root_base / f"simulation_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        try:
            simulation_root.mkdir(parents=True, exist_ok=True)
        except OSError:
            logger.exception("No se pudo crear la carpeta de simulación: %s", simulation_root)
            summary.simulation_root = ""
            return summary

        for item in summary.items:
            planned_path = Path(item.planned_path) if item.planned_path else None
            source_path = Path(item.source_pdf_path) if item.source_pdf_path else None
            if item.status != DossierStatus.PLANNED or not planned_path or not source_path.is_file():
                continue
            destination = self._map_destination(root_path, planned_path, simulation_root, item.row_number, item.cp, item.serie)
            try:
                self._copy_with_long_path_support(source_path, destination)
                item.simulation_path = str(destination)
            except OSError:
                logger.exception("No se pudo materializar la simulación para %s -> %s", source_path, destination)
                item.simulation_path = ""

        summary.simulation_root = str(simulation_root)
        return summary

    @staticmethod
    def _map_destination(root_path: Path, planned_path: Path, simulation_root: Path, row_number: int, cp: str, serie: str) -> Path:
        try:
            relative_path = planned_path.relative_to(root_path)
            return simulation_root / relative_path
        except Exception:
            cp_folder = cp or f"fila_{row_number}"
            return simulation_root / cp_folder / serie / planned_path.name

    def _copy_with_long_path_support(self, source_path: Path, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        source_for_copy = self._with_windows_long_path_prefix(source_path)
        destination_for_copy = self._with_windows_long_path_prefix(destination)
        destination_parent_for_copy = self._with_windows_long_path_prefix(destination.parent)
        os.makedirs(destination_parent_for_copy, exist_ok=True)
        shutil.copy2(source_for_copy, destination_for_copy)

    @staticmethod
    def _with_windows_long_path_prefix(path: Path) -> str:
        raw = str(path.resolve())
        if os.name != 'nt':
            return raw
        slash = "\\"
        long_prefix = slash * 2 + "?" + slash
        if raw.startswith(long_prefix):
            return raw
        if raw.startswith(slash * 2):
            return long_prefix + "UNC" + slash + raw.lstrip(slash)
        return long_prefix + raw
