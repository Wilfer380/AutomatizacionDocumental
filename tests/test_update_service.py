from __future__ import annotations

import json
from pathlib import Path

from word_excel_pdf_automation.services.update_service import UpdateService


def test_compare_versions_handles_semver_like_numbers() -> None:
    service = UpdateService()

    assert service.compare_versions("1.0.1", "1.0.0") == 1
    assert service.compare_versions("1.0.0", "1.0.0") == 0
    assert service.compare_versions("1.0.0", "1.0.1") == -1
    assert service.compare_versions("1.10.0", "1.2.9") == 1


def test_check_for_update_returns_available_release(tmp_path: Path, monkeypatch) -> None:
    shared_root = tmp_path / "updates"
    shared_root.mkdir()
    installer_path = shared_root / "AutomatizacionDocumentalSetup_1.0.1.exe"
    installer_path.write_text("binary", encoding="utf-8")
    (shared_root / "latest.json").write_text(
        json.dumps(
            {
                "version": "1.0.1",
                "installer_name": installer_path.name,
                "notes": "Correcciones",
                "published_at": "2026-07-14T10:00:00",
            }
        ),
        encoding="utf-8",
    )

    settings_path = tmp_path / "update_settings.json"
    settings_path.write_text(
        json.dumps(
            {
                "enabled": True,
                "shared_root": str(shared_root),
                "manifest_name": "latest.json",
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setenv("AUTOMATIZACION_DOCUMENTAL_UPDATE_CONFIG", str(settings_path))

    result = UpdateService().check_for_update("1.0.0")

    assert result.checked is True
    assert result.available is True
    assert result.release is not None
    assert result.release.version == "1.0.1"
    assert result.release.installer_path == installer_path


def test_check_for_update_skips_when_disabled(tmp_path: Path, monkeypatch) -> None:
    settings_path = tmp_path / "update_settings.json"
    settings_path.write_text(json.dumps({"enabled": False}), encoding="utf-8")
    monkeypatch.setenv("AUTOMATIZACION_DOCUMENTAL_UPDATE_CONFIG", str(settings_path))

    result = UpdateService().check_for_update("1.0.0")

    assert result.checked is False
    assert result.available is False
    assert "deshabilitadas" in result.reason.lower()
