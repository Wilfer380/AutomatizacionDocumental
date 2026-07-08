from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from word_excel_pdf_automation.dossier_models import DossierConfig, DossierExecutionMode, DossierStatus
from word_excel_pdf_automation.services.dossier_distribution_service import DossierDistributionService, RealModeConfirmationRequiredError
from word_excel_pdf_automation.services.dossier_service import DossierService
from word_excel_pdf_automation.services.dossier_validator_service import DossierValidatorService


def _write_workbook(path: Path, headers: list[str], rows: list[list[object]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Hoja1"
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


def _build_config(*, root_path: Path, excel_path: Path, simulation_only: bool = True) -> DossierConfig:
    source_root = excel_path.parent
    return DossierConfig.from_mapping(
        {
            "root_path": str(root_path),
            "excel_path": str(excel_path),
            "dossier_folder_name": "06_DOSSIER",
            "simulation_only": simulation_only,
            "cp_synonyms": ["Centro Proyecto"],
            "serie_synonyms": ["Número de serie"],
            "pdf_sources": [
                {"document_name": "Descriptivo de pintura", "source_pdf_path": str(source_root / "sources" / "5.2 Descriptivo de pintura.pdf"), "target_folder": "5 Pintura", "final_name_pattern": "5.2 Descriptivo de pintura.pdf", "mandatory": True},
                {"document_name": "Certificado trazabilidad", "source_pdf_path": str(source_root / "sources" / "6.1 Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf"), "target_folder": "6 Trazabilidad", "final_name_pattern": "6.1 Certificado de Trazabilidad de Ensayos de Pintura SN {serie}.pdf", "mandatory": True},
                {"document_name": "Resultados adherencia", "source_pdf_path": str(source_root / "sources" / "7.1 Comunicado técnico - Resultados.pdf"), "target_folder": "7 Ensayos", "final_name_pattern": "7.1 Comunicado técnico - Resultados de adherencia.pdf", "mandatory": True},
                {"document_name": "Informe laboratorio", "source_pdf_path": str(source_root / "sources" / "7.2 Informe de Ensayo Laboratorio.pdf"), "target_folder": "7 Ensayos", "final_name_pattern": "7.2 Informe de Ensayo Laboratorio - Prueba adherencia.pdf", "mandatory": True},
            ],
        },
        source_path=excel_path.with_suffix(".json"),
    )


def _build_cp_tree(root_path: Path, cp_name: str, *series: str) -> Path:
    cp_folder = root_path / cp_name
    dossier = cp_folder / "06_DOSSIER"
    planos = dossier / "2 Planos"
    (dossier / "5 Procedimiento de fabricación").mkdir(parents=True, exist_ok=True)
    (dossier / "6 Trazabilidad").mkdir(parents=True, exist_ok=True)
    (dossier / "7 Ensayos").mkdir(parents=True, exist_ok=True)
    planos.mkdir(parents=True, exist_ok=True)
    for index, serie in enumerate(series, start=1):
        (planos / f"2.3.{index} PLACA ID {serie}.pdf").write_bytes(b"%PDF-1.4\n")
    return cp_folder


def _create_pdf(path: Path, content: bytes = b"%PDF-1.4\n% test\n") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return path


def _create_all_sources(root: Path) -> None:
    _create_pdf(root / "sources" / "5.2 Descriptivo de pintura.pdf")
    _create_pdf(root / "sources" / "6.1 Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf")
    _create_pdf(root / "sources" / "7.1 Comunicado técnico - Resultados.pdf")
    _create_pdf(root / "sources" / "7.2 Informe de Ensayo Laboratorio.pdf")


def test_headers_and_missing_header_detection(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "headers.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    config = _build_config(root_path=tmp_path, excel_path=workbook_path)
    workbook_info = DossierValidatorService().inspect_workbook(config)
    assert workbook_info.cp_column == "Centro Proyecto"
    missing_path = _write_workbook(tmp_path / "missing.xlsx", ["Foo", "Bar"], [["CP-01", "x"]])
    summary = DossierService().run_simulation(_build_config(root_path=tmp_path / "missing-root", excel_path=missing_path), report_dir=tmp_path / "reports")
    assert summary.blocked == 1


def test_series_is_found_in_planos_file_names(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "series.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)
    assert validated[0].status == DossierStatus.VALID
    assert validated[0].series_in_planos is True


def test_multiple_cp_candidates_with_same_series_are_all_kept(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "multi.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-02", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-02 A", "S-001")
    _build_cp_tree(root_path, "CP-02 B", "S-001")
    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)
    assert validated[0].status == DossierStatus.VALID
    assert len(validated[0].matched_dossier_folders) == 2


def test_simulation_does_not_write_files_and_reports_progress(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "simulation.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    updates: list[tuple[int, int, str]] = []
    summary = DossierService().run_simulation(_build_config(root_path=root_path, excel_path=workbook_path), report_dir=tmp_path / "reports", progress_callback=lambda c, t, m: updates.append((c, t, m)))
    planned = next(item for item in summary.items if item.rule_name == "Descriptivo de pintura")
    assert summary.execution_mode == DossierExecutionMode.SIMULATION
    assert Path(planned.planned_path).exists() is False
    assert updates[-1][1] == summary.planned_actions


def test_missing_sources_respect_mandatory_flag(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "sources.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_pdf(tmp_path / "sources" / "5.2 Descriptivo de pintura.pdf")
    _create_pdf(tmp_path / "sources" / "7.2 Informe de Ensayo Laboratorio.pdf")
    config = _build_config(root_path=root_path, excel_path=workbook_path)
    config.pdf_sources[1].source_pdf_path = tmp_path / "sources" / "missing.pdf"
    config.pdf_sources[2].mandatory = False
    config.pdf_sources[2].source_pdf_path = tmp_path / "sources" / "missing-optional.pdf"
    summary = DossierService().run_simulation(config, report_dir=tmp_path / "reports-sources")
    statuses = [item.status for item in summary.items if item.rule_name != "validation"]
    assert DossierStatus.PLANNED in statuses
    assert DossierStatus.ERROR in statuses
    assert DossierStatus.SKIPPED in statuses


def test_real_copy_creates_backup_before_replace(tmp_path: Path) -> None:
    source_pdf = _create_pdf(tmp_path / "source.pdf", b"source")
    destination = tmp_path / "destination.pdf"
    destination.write_bytes(b"old")
    result = DossierDistributionService().copy_or_replace(source_pdf, destination, confirm_real=True, backup_root=tmp_path / "backups")
    assert result.status == DossierStatus.REPLACED
    assert Path(result.backup_path).exists()


def test_real_mode_requires_confirmation(tmp_path: Path) -> None:
    with pytest.raises(RealModeConfirmationRequiredError):
        DossierDistributionService().copy_or_replace(_create_pdf(tmp_path / "source.pdf"), tmp_path / "destination.pdf", confirm_real=False)


def test_phase1_routing_success_and_mismatch(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "phase1.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    source_pdf = _create_pdf(tmp_path / "sources" / "Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf")
    service = DossierService()
    summary_ok = service.run_distribution(_build_config(root_path=root_path, excel_path=workbook_path), confirm_real=False, report_dir=tmp_path / "reports-ok", phase1_items=[SimpleNamespace(row_number=2, series="S-001", pdf_filename=source_pdf.name, pdf_path=str(source_pdf))])
    summary_bad = service.run_distribution(_build_config(root_path=root_path, excel_path=workbook_path), confirm_real=False, report_dir=tmp_path / "reports-bad", phase1_items=[SimpleNamespace(row_number=2, series="S-999", pdf_filename="bad.pdf", pdf_path=str(source_pdf))])
    assert next(item for item in summary_ok.items if item.rule_name == "phase1-routing").status == DossierStatus.PLANNED
    assert next(item for item in summary_bad.items if item.rule_name == "phase1-routing").status == DossierStatus.SKIPPED


def test_distribution_plans_multiple_targets_for_same_cp(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "multi-plan.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-02", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-02 A", "S-001")
    _build_cp_tree(root_path, "CP-02 B", "S-001")
    _create_all_sources(tmp_path)
    summary = DossierService().run_simulation(_build_config(root_path=root_path, excel_path=workbook_path), report_dir=tmp_path / "reports-multi")
    descriptivos = [item for item in summary.items if item.rule_name == "Descriptivo de pintura"]
    assert len(descriptivos) == 2
