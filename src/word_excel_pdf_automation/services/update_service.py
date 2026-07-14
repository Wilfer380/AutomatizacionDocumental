from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from ..config import APP_NAME, read_json_file
from ..update_models import UpdateCheckResult, UpdateRelease, UpdateSettings


class UpdateService:
    def __init__(self, settings_file_name: str = "update_settings.json") -> None:
        self.settings_file_name = settings_file_name

    def load_settings(self) -> UpdateSettings:
        settings_path = self._resolve_settings_path()
        if settings_path is None or not settings_path.is_file():
            return UpdateSettings()

        payload = read_json_file(settings_path)
        shared_root_value = str(payload.get("shared_root", "")).strip()
        shared_root = Path(shared_root_value) if shared_root_value else None
        return UpdateSettings(
            enabled=bool(payload.get("enabled", False)),
            shared_root=shared_root,
            manifest_name=str(payload.get("manifest_name", "latest.json")).strip() or "latest.json",
            installer_name=str(payload.get("installer_name", "")).strip(),
            channel=str(payload.get("channel", "stable")).strip() or "stable",
        )

    def check_for_update(self, current_version: str) -> UpdateCheckResult:
        settings = self.load_settings()
        if not settings.enabled:
            return UpdateCheckResult(checked=False, available=False, current_version=current_version, reason="Actualizaciones deshabilitadas.")
        if settings.shared_root is None:
            return UpdateCheckResult(checked=False, available=False, current_version=current_version, reason="No hay carpeta compartida configurada.")
        if not settings.shared_root.exists():
            return UpdateCheckResult(checked=False, available=False, current_version=current_version, reason=f"No existe la carpeta compartida: {settings.shared_root}")

        manifest_path = settings.shared_root / settings.manifest_name
        if not manifest_path.is_file():
            return UpdateCheckResult(checked=False, available=False, current_version=current_version, reason=f"No existe el manifiesto: {manifest_path}")

        payload = read_json_file(manifest_path)
        version = str(payload.get("version", "")).strip()
        if not version:
            return UpdateCheckResult(checked=False, available=False, current_version=current_version, reason="El manifiesto no trae versi?n.")

        installer_name = str(payload.get("installer_name", settings.installer_name)).strip()
        if not installer_name:
            return UpdateCheckResult(checked=False, available=False, current_version=current_version, reason="El manifiesto no trae instalador.")

        installer_path = settings.shared_root / installer_name
        if not installer_path.is_file():
            return UpdateCheckResult(
                checked=False,
                available=False,
                current_version=current_version,
                reason=f"No existe el instalador publicado: {installer_path}",
            )

        if self.compare_versions(version, current_version) <= 0:
            return UpdateCheckResult(checked=True, available=False, current_version=current_version, reason="La versi?n actual ya est? al d?a.")

        release = UpdateRelease(
            version=version,
            installer_path=installer_path,
            notes=str(payload.get("notes", "")).strip(),
            published_at=str(payload.get("published_at", "")).strip(),
            mandatory=bool(payload.get("mandatory", False)),
        )
        return UpdateCheckResult(checked=True, available=True, current_version=current_version, release=release)

    def stage_installer(self, release: UpdateRelease) -> Path:
        if not release.installer_path.is_file():
            raise FileNotFoundError(f"No existe el instalador publicado: {release.installer_path}")

        target_dir = Path(tempfile.gettempdir()) / APP_NAME / "updates" / release.version
        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = target_dir / release.installer_path.name
        shutil.copy2(release.installer_path, target_path)
        return target_path

    def launch_installer(self, installer_path: Path) -> None:
        if not installer_path.is_file():
            raise FileNotFoundError(f"No existe el instalador preparado: {installer_path}")

        if hasattr(os, "startfile"):
            os.startfile(str(installer_path))  # type: ignore[attr-defined]
            return
        subprocess.Popen([str(installer_path)], close_fds=True)

    def compare_versions(self, left: str, right: str) -> int:
        left_parts = self._normalize_version(left)
        right_parts = self._normalize_version(right)
        size = max(len(left_parts), len(right_parts))
        padded_left = left_parts + (0,) * (size - len(left_parts))
        padded_right = right_parts + (0,) * (size - len(right_parts))
        if padded_left > padded_right:
            return 1
        if padded_left < padded_right:
            return -1
        return 0

    def _normalize_version(self, value: str) -> tuple[int, ...]:
        clean = value.strip().replace("-", ".").replace("_", ".")
        parts: list[int] = []
        for raw_part in clean.split("."):
            digits = "".join(character for character in raw_part if character.isdigit())
            parts.append(int(digits or "0"))
        return tuple(parts) if parts else (0,)

    def _resolve_settings_path(self) -> Path | None:
        env_path = os.getenv("AUTOMATIZACION_DOCUMENTAL_UPDATE_CONFIG", "").strip()
        if env_path:
            candidate = Path(env_path).expanduser()
            if candidate.exists():
                return candidate

        for directory in self._candidate_setting_directories():
            candidate = directory / self.settings_file_name
            if candidate.exists():
                return candidate
        return None

    def _candidate_setting_directories(self) -> list[Path]:
        directories: list[Path] = []

        bundle_dir = Path(getattr(sys, "_MEIPASS", "")).resolve() if getattr(sys, "_MEIPASS", "") else None
        if bundle_dir:
            directories.append(bundle_dir)

        executable_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else None
        if executable_dir:
            directories.append(executable_dir)

        project_root = Path(__file__).resolve().parents[3]
        directories.append(project_root)

        program_data = os.getenv("PROGRAMDATA", "").strip()
        if program_data:
            directories.append(Path(program_data) / APP_NAME)

        app_data = os.getenv("APPDATA", "").strip()
        if app_data:
            directories.append(Path(app_data) / APP_NAME)

        unique: list[Path] = []
        seen: set[str] = set()
        for directory in directories:
            key = str(directory).lower()
            if key not in seen:
                seen.add(key)
                unique.append(directory)
        return unique
