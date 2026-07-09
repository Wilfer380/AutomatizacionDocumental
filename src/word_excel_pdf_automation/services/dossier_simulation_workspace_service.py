from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from ..dossier_models import DossierExecutionMode, DossierRunSummary, DossierStatus


class DossierSimulationWorkspaceService:
    def materialize(self, root_path: Path, summary: DossierRunSummary, simulation_root_base: Path) -> DossierRunSummary:
        if summary.execution_mode != DossierExecutionMode.SIMULATION:
            return summary

        simulation_root = simulation_root_base / f"simulation_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        simulation_root.mkdir(parents=True, exist_ok=True)
        for item in summary.items:
            planned_path = Path(item.planned_path) if item.planned_path else None
            source_path = Path(item.source_pdf_path) if item.source_pdf_path else None
            if item.status != DossierStatus.PLANNED or not planned_path or not source_path.exists():
                continue
            destination = self._map_destination(root_path, planned_path, simulation_root, item.row_number, item.cp, item.serie)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, destination)
            item.simulation_path = str(destination)

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
