from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import shutil


@dataclass(slots=True)
class BackupPlan:
    source_path: Path
    destination_path: Path
    simulated: bool = True
    status: str = "planned"
    observation: str = ""


class DossierBackupService:
    def plan_backup(self, source_path: Path | None, backup_root: Path, simulated: bool = True) -> BackupPlan:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destination_path = backup_root / f"config_backup_{timestamp}.json"
        if source_path is None:
            return BackupPlan(source_path=Path(), destination_path=destination_path, simulated=simulated, status="skipped", observation="No se proporcionó un archivo de configuración.")
        if not source_path.exists():
            return BackupPlan(source_path=source_path, destination_path=destination_path, simulated=simulated, status="missing", observation="El archivo de configuración de origen no existe.")
        if simulated:
            return BackupPlan(source_path=source_path, destination_path=destination_path, simulated=True, status="planned", observation="Solo simulación; no se copió ningún archivo.")

        destination_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination_path)
        return BackupPlan(source_path=source_path, destination_path=destination_path, simulated=False, status="copied", observation="Respaldo creado.")

    def plan_destination_backup(self, destination_path: Path | None, backup_root: Path, simulated: bool = True) -> BackupPlan:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if destination_path is None:
            destination_path = backup_root / f"destination_backup_{timestamp}.bak"
            return BackupPlan(source_path=Path(), destination_path=destination_path, simulated=simulated, status="skipped", observation="No se proporcionó un archivo de destino.")

        backup_path = backup_root / f"destination_backup_{destination_path.stem}_{timestamp}{destination_path.suffix}"
        if not destination_path.exists():
            return BackupPlan(source_path=destination_path, destination_path=backup_path, simulated=simulated, status="missing", observation="El archivo de destino no existe.")
        if simulated:
            return BackupPlan(source_path=destination_path, destination_path=backup_path, simulated=True, status="planned", observation="Solo simulación; el respaldo del destino no se copió.")

        backup_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(destination_path, backup_path)
        return BackupPlan(source_path=destination_path, destination_path=backup_path, simulated=False, status="copied", observation="Respaldo del destino creado.")
