from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook

from word_excel_pdf_automation.dossier_models import DossierConfig, DossierExecutionMode, DossierStatus
from word_excel_pdf_automation.services.dossier_distribution_service import DossierDistributionService, RealModeConfirmationRequiredError
from word_excel_pdf_automation.services.dossier_service import DossierService
from word_excel_pdf_automation.services.dossier_simulation_workspace_service import DossierSimulationWorkspaceService
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


def _build_config(*, root_path: Path, excel_path: Path, simulation_only: bool = True, cp_filter: str = '', serie_filter: str = '', cp_filters: list[str] | None = None, serie_filters: list[str] | None = None, replace_existing: bool = True) -> DossierConfig:
    source_root = excel_path.parent
    return DossierConfig.from_mapping(
        {
            "root_path": str(root_path),
            "excel_path": str(excel_path),
            "dossier_folder_name": "06_DOSSIER",
            "simulation_only": simulation_only,
            "replace_existing": replace_existing,
            "cp_filter": cp_filter,
            "serie_filter": serie_filter,
            "cp_filters": cp_filters or [],
            "serie_filters": serie_filters or [],
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


def _build_nested_dossier_tree(root_path: Path, cp_name: str, nested_name: str, dossier_name: str, *series: str) -> tuple[Path, Path]:
    cp_folder = root_path / cp_name
    dossier_parent = cp_folder if nested_name in {"", "."} else cp_folder / nested_name
    dossier = dossier_parent / dossier_name
    planos = dossier / "2 Planos"
    (dossier / "5 Procedimiento de fabricacion").mkdir(parents=True, exist_ok=True)
    (dossier / "6 Registros Informes de Inspeccion").mkdir(parents=True, exist_ok=True)
    (dossier / "7 Pruebas Electricas").mkdir(parents=True, exist_ok=True)
    planos.mkdir(parents=True, exist_ok=True)
    for index, serie in enumerate(series, start=1):
        (planos / f"2.3.{index} PLACA ID {serie}.pdf").write_bytes(b"%PDF-1.4\n")
    return cp_folder, dossier


def _build_legacy_distribution_root(root_path: Path, cp_name: str, branch_name: str, *series: str) -> tuple[Path, Path]:
    cp_folder = root_path / cp_name
    dossier = cp_folder / branch_name
    planos = dossier / "2 Planos"
    (dossier / "5 Procedimiento de fabricacion").mkdir(parents=True, exist_ok=True)
    (dossier / "6 Registros Informes de Inspeccion").mkdir(parents=True, exist_ok=True)
    (dossier / "7 Pruebas Electricas").mkdir(parents=True, exist_ok=True)
    planos.mkdir(parents=True, exist_ok=True)
    for index, serie in enumerate(series, start=1):
        (planos / f"2.3.{index} PLACA ID {serie}.pdf").write_bytes(b"%PDF-1.4\n")
    return cp_folder, dossier


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


def test_series_is_found_in_folder_6_when_missing_from_planos(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "series-folder6.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01")
    folder_6 = cp_folder / "06_DOSSIER" / "6 Registros de informe de inspeccion"
    (cp_folder / "06_DOSSIER" / "6 Trazabilidad").rename(folder_6)
    _create_pdf(folder_6 / "registro_S-001.pdf")
    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)
    assert validated[0].status == DossierStatus.VALID
    assert "6 Registros de informe" in validated[0].folder_6



def test_series_is_found_when_dossier_is_nested_inside_cp_folder(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "nested-dossier.xlsx", ["Centro Proyecto", "Numero de serie"], [["15915431", "1059083746"]])
    root_path = tmp_path / "root"
    _build_nested_dossier_tree(
        root_path,
        "15915431_300 kVA_0.48kV_SLA COL_OV51043843",
        "OV 51002644-OC 4500004252_300kVA_15_D",
        "06_DOSSIER",
        "1059083746",
    )

    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)

    assert validated[0].status == DossierStatus.VALID
    assert "OV51043843" in validated[0].cp_folder
    assert validated[0].dossier_folder.endswith("06_DOSSIER")



def test_series_is_found_in_legacy_structure_without_06_dossier(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "legacy-structure.xlsx", ["Centro Proyecto", "Numero de serie"], [["15915431", "1059083746"]])
    root_path = tmp_path / "root"
    _build_legacy_distribution_root(
        root_path,
        "15915431_300 kVA_0.48kV_SLA COL_OV51043843",
        "OV 51002644-OC 4500004252_300kVA_15_D",
        "1059083746",
    )

    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)

    assert validated[0].status == DossierStatus.VALID
    assert validated[0].dossier_folder.endswith("OV 51002644-OC 4500004252_300kVA_15_D")
    assert "2 Planos" in validated[0].planos_folder


def test_series_is_found_when_dossier_name_has_suffix(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "dossier-suffix.xlsx", ["Centro Proyecto", "Numero de serie"], [["15776551", "1084638314"]])
    root_path = tmp_path / "root"
    _cp_folder, dossier = _build_nested_dossier_tree(
        root_path,
        "15776551_SLA COL_4.2kV_150 kVA_OV52128104",
        ".",
        "06_DOSSIER (POSICION 10-20)",
        "1084638314",
    )
    folder_6 = dossier / "6 Registros Informes de Inspeccion"
    _create_pdf(folder_6 / "6.1.8 SN 1084638314.pdf")

    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)

    assert validated[0].status == DossierStatus.VALID
    assert validated[0].cp_folder.endswith("OV52128104")
    assert "POSICION 10-20" in validated[0].dossier_folder


def test_series_is_found_anywhere_inside_dossier_for_same_cp(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "series-anywhere.xlsx", ["Centro Proyecto", "Numero de serie"], [["15915431", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "15915431_300KVA_4.2kV_SLA COL_OV52618579")
    planos = cp_folder / "06_DOSSIER" / "2 Planos"
    for item in planos.iterdir():
        item.unlink()
    folder_6 = cp_folder / "06_DOSSIER" / "6 Trazabilidad"
    for item in folder_6.iterdir():
        if item.is_file():
            item.unlink()
    folder_8 = cp_folder / "06_DOSSIER" / "8 Cert de Calidad de Materiales"
    folder_8.mkdir(parents=True, exist_ok=True)
    _create_pdf(folder_8 / "8.1 Cert Fluido Aislante S-001.pdf")

    _, rows = DossierValidatorService().load_rows(_build_config(root_path=root_path, excel_path=workbook_path))
    validated = DossierValidatorService().validate_paths(_build_config(root_path=root_path, excel_path=workbook_path), rows)

    assert validated[0].status == DossierStatus.VALID
    assert validated[0].cp_folder.endswith("OV52618579")


def test_cp_code_is_resolved_from_numeric_folder_names(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "numeric-cp.xlsx", ["Centro Proyecto", "Numero de serie"], [["15915431", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "15915431_300KVA_4.2kV_SLA COL_OV52618579", "S-001")

    service = DossierValidatorService()
    matches = service.find_cp_folders(root_path, "15915431")

    assert len(matches) == 1
    assert matches[0].name.startswith("15915431_")

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
    phase1_ok = next(item for item in summary_ok.items if item.rule_name == "phase1-routing")
    assert phase1_ok.status == DossierStatus.PLANNED
    assert "6 Trazabilidad" in phase1_ok.target_folder
    assert Path(phase1_ok.planned_path).name.startswith("6.")
    assert "Certificado de Trazabilidad de Ensayos de Pintura SN S-001" in Path(phase1_ok.planned_path).name
    assert next(item for item in summary_bad.items if item.rule_name == "phase1-routing").status == DossierStatus.SKIPPED


def test_phase1_routing_finds_folder_6_from_dossier_targets(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "phase1-folder6-derived.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    source_pdf = _create_pdf(tmp_path / "sources" / "Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf")
    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path),
        confirm_real=False,
        report_dir=tmp_path / "reports-folder6-derived",
        phase1_items=[SimpleNamespace(row_number=2, series="S-001", pdf_filename=source_pdf.name, pdf_path=str(source_pdf))],
    )
    phase1_item = next(item for item in summary.items if item.rule_name == "phase1-routing")
    assert phase1_item.status == DossierStatus.PLANNED
    assert phase1_item.target_folder.endswith("6 Trazabilidad")
    assert Path(phase1_item.planned_path).name.startswith("6.")
    assert source_pdf.stem in Path(phase1_item.planned_path).stem


def test_phase1_routing_targets_folder_6_for_matching_series(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "phase1-folder6.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    source_pdf = _create_pdf(tmp_path / "sources" / "Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf")
    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path),
        confirm_real=False,
        report_dir=tmp_path / "reports-folder6",
        phase1_items=[SimpleNamespace(row_number=2, series="S-001", pdf_filename=source_pdf.name, pdf_path=str(source_pdf))],
    )
    phase1_item = next(item for item in summary.items if item.rule_name == "phase1-routing")
    assert phase1_item.status == DossierStatus.PLANNED
    assert "6 Trazabilidad" in phase1_item.target_folder
    assert "6 Trazabilidad" in phase1_item.planned_path
    assert Path(phase1_item.planned_path).name.startswith("6.")


def test_phase1_routing_uses_folder6_escalability_per_series(tmp_path: Path) -> None:
    workbook_path = _write_workbook(
        tmp_path / "phase1-folder6-escalability.xlsx",
        ["Centro Proyecto", "Numero de serie"],
        [["CP-01", "100001"], ["CP-01", "100002"], ["CP-01", "100003"]],
    )
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01", "100001", "100002", "100003")
    _create_all_sources(tmp_path)
    folder_6 = cp_folder / "06_DOSSIER" / "6 Trazabilidad"
    _create_pdf(folder_6 / "6.1.1 Registro previo SN 100001.pdf")
    _create_pdf(folder_6 / "6.1.2 Registro previo SN 100002.pdf")
    _create_pdf(folder_6 / "6.1.3 Registro previo SN 100003.pdf")

    phase1_items = [
        SimpleNamespace(row_number=2, series="100001", pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN 100001.pdf", pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-100001.pdf"))),
        SimpleNamespace(row_number=3, series="100002", pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN 100002.pdf", pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-100002.pdf"))),
        SimpleNamespace(row_number=4, series="100003", pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN 100003.pdf", pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-100003.pdf"))),
    ]

    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path),
        confirm_real=False,
        report_dir=tmp_path / "reports-folder6-escalability",
        phase1_items=phase1_items,
    )

    phase1_routes = [item for item in summary.items if item.rule_name == "phase1-routing"]
    names = [Path(item.planned_path).name for item in phase1_routes]

    assert names == [
        "6.2.1 Certificado de Trazabilidad de Ensayos de Pintura SN 100001.pdf",
        "6.2.2 Certificado de Trazabilidad de Ensayos de Pintura SN 100002.pdf",
        "6.2.3 Certificado de Trazabilidad de Ensayos de Pintura SN 100003.pdf",
    ]


def test_simulation_skips_existing_document_when_replace_disabled(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "existing-skip.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    config = _build_config(root_path=root_path, excel_path=workbook_path, replace_existing=False)
    config.pdf_sources[2].final_name_pattern = "7.1 Comunicado tecnico - Resultados de adherencia.pdf"
    existing_target = cp_folder / "06_DOSSIER" / "7 Ensayos" / config.pdf_sources[2].final_name_pattern
    _create_pdf(existing_target)

    summary = DossierService().run_simulation(
        config,
        report_dir=tmp_path / "reports-existing-skip",
    )

    comunicado = next(item for item in summary.items if item.rule_name == "Resultados adherencia")
    assert comunicado.status == DossierStatus.SKIPPED
    assert "ya" in comunicado.observation.lower()


def test_real_distribution_replaces_legacy_folder5_document(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "replace-folder5.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    legacy_path = cp_folder / "06_DOSSIER" / "5 Procedimiento de fabricacion" / "5.2 Procedimiento Aplicacion Pintura Tanques 2023.03.08.pdf"
    _create_pdf(legacy_path, b"%PDF-1.4\nold\n")

    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path, simulation_only=False, replace_existing=True),
        confirm_real=True,
        report_dir=tmp_path / "reports-replace-folder5",
    )

    descriptivo = next(item for item in summary.items if item.rule_name == "Descriptivo de pintura")
    new_path = cp_folder / "06_DOSSIER" / "5 Procedimiento de fabricacion" / "5.2 Descriptivo de pintura.pdf"
    assert descriptivo.status == DossierStatus.REPLACED
    assert new_path.exists()
    assert not legacy_path.exists()


def test_real_distribution_deduplicates_shared_folder5_and_folder7_documents(tmp_path: Path) -> None:
    workbook_path = _write_workbook(
        tmp_path / "shared-docs.xlsx",
        ["Centro Proyecto", "Numero de serie"],
        [["CP-01", "S-001"], ["CP-01", "S-002"], ["CP-01", "S-003"]],
    )
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001", "S-002", "S-003")
    _create_all_sources(tmp_path)

    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path, simulation_only=False, replace_existing=True),
        confirm_real=True,
        report_dir=tmp_path / "reports-shared-docs",
    )

    folder5_items = [item for item in summary.items if item.rule_name == "Descriptivo de pintura"]
    folder7_items = [item for item in summary.items if item.rule_name in {"Resultados adherencia", "Informe laboratorio"}]
    folder6_items = [item for item in summary.items if item.rule_name == "Certificado trazabilidad"]

    assert len(folder5_items) == 1
    assert len(folder7_items) == 2
    assert len(folder6_items) == 3
    assert all(item.status in {DossierStatus.COPIED, DossierStatus.REPLACED} for item in folder5_items + folder7_items + folder6_items)
    assert summary.warnings == 0


def test_real_distribution_skips_existing_folder6_certificate_even_if_section_changes(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "folder6-existing.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)

    folder_6 = cp_folder / "06_DOSSIER" / "6 Trazabilidad"
    _create_pdf(folder_6 / "6.11.10 Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf", b"%PDF-1.4\nexisting\n")

    phase1_item = SimpleNamespace(
        row_number=2,
        series="S-001",
        pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf",
        pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-s001.pdf")),
    )

    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path, simulation_only=False, replace_existing=True),
        confirm_real=True,
        report_dir=tmp_path / "reports-folder6-existing",
        phase1_items=[phase1_item],
    )

    phase1_result = next(item for item in summary.items if item.rule_name == "phase1-routing")
    certificates = sorted(path.name for path in folder_6.glob('*Certificado de Trazabilidad*.pdf'))

    assert phase1_result.status == DossierStatus.SKIPPED
    assert len(certificates) == 1
    assert certificates[0] == "6.11.10 Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf"


def test_phase1_routing_keeps_all_series_when_folder5_and_folder7_are_deduplicated(tmp_path: Path) -> None:
    workbook_path = _write_workbook(
        tmp_path / "phase1-shared-docs.xlsx",
        ["Centro Proyecto", "Numero de serie"],
        [["CP-01", "S-001"], ["CP-01", "S-002"], ["CP-01", "S-003"]],
    )
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001", "S-002", "S-003")
    _create_all_sources(tmp_path)

    phase1_items = [
        SimpleNamespace(row_number=2, series="S-001", pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN S-001.pdf", pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-s001.pdf"))),
        SimpleNamespace(row_number=3, series="S-002", pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN S-002.pdf", pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-s002.pdf"))),
        SimpleNamespace(row_number=4, series="S-003", pdf_filename="Certificado de Trazabilidad de Ensayos de Pintura SN S-003.pdf", pdf_path=str(_create_pdf(tmp_path / "sources" / "phase1-s003.pdf"))),
    ]

    summary = DossierService().run_distribution(
        _build_config(root_path=root_path, excel_path=workbook_path, simulation_only=False, replace_existing=True),
        confirm_real=True,
        report_dir=tmp_path / "reports-phase1-shared-docs",
        phase1_items=phase1_items,
    )

    phase1_routes = [item for item in summary.items if item.rule_name == "phase1-routing"]
    assert len(phase1_routes) == 3
    assert {item.serie for item in phase1_routes} == {"S-001", "S-002", "S-003"}
    assert all(item.status in {DossierStatus.COPIED, DossierStatus.REPLACED, DossierStatus.SKIPPED} for item in phase1_routes)
    assert all("No se encontr" not in (item.skipped_reason or "") for item in phase1_routes)


def test_distribution_plans_multiple_targets_for_same_cp(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "multi-plan.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-02", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-02 A", "S-001")
    _build_cp_tree(root_path, "CP-02 B", "S-001")
    _create_all_sources(tmp_path)
    summary = DossierService().run_simulation(_build_config(root_path=root_path, excel_path=workbook_path), report_dir=tmp_path / "reports-multi")
    descriptivos = [item for item in summary.items if item.rule_name == "Descriptivo de pintura"]
    assert len(descriptivos) == 2




def test_simulation_filters_rows_by_cp_and_serie(tmp_path: Path) -> None:
    workbook_path = _write_workbook(
        tmp_path / 'filter.xlsx',
        ['Centro Proyecto', 'Numero de serie'],
        [['CP-01', 'S-001'], ['CP-02', 'S-002']],
    )
    root_path = tmp_path / 'root'
    _build_cp_tree(root_path, 'CP-01', 'S-001')
    _build_cp_tree(root_path, 'CP-02', 'S-002')
    _create_all_sources(tmp_path)

    summary = DossierService().run_simulation(
        _build_config(root_path=root_path, excel_path=workbook_path, cp_filter='CP-02', serie_filter='S-002'),
        report_dir=tmp_path / 'reports-filter',
    )

    assert summary.total_rows == 1
    assert all(item.cp == 'CP-02' for item in summary.items)
    assert all(item.serie == 'S-002' for item in summary.items)

def test_simulation_creates_real_preview_workspace(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "simulation-workspace.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    summary = DossierService().run_simulation(_build_config(root_path=root_path, excel_path=workbook_path), report_dir=tmp_path / "reports-sim-workspace")
    descriptivo = next(item for item in summary.items if item.rule_name == "Descriptivo de pintura")
    assert summary.simulation_root
    assert Path(summary.simulation_root).exists()
    assert descriptivo.simulation_path
    assert Path(descriptivo.simulation_path).exists()
    assert Path(descriptivo.planned_path).exists() is False



def test_simulation_preserves_original_relative_structure(tmp_path: Path) -> None:
    service = DossierSimulationWorkspaceService()
    simulation_root = tmp_path / "simulation_root"
    root_path = Path(r"Q:\GROUPS\CO_MDE_DISENO_DI\01_ORDERS\02_DOCUMENTS_APPROVAL_CERTIFIED")
    planned_path = root_path / r"15915431_300KVA_SLA COL_OV51396245\06_DOSSIER_OV51396245-Pos.350-370-OC 4500005244_300kVA_54_D(carpeta 2)\5 Procedimiento de Fabricacion\5.2 Descriptivo de pintura.pdf"

    destination = service._map_destination(root_path, planned_path, simulation_root, 33, "15915431", "1066388130")

    assert destination == simulation_root / planned_path.relative_to(root_path)



def test_windows_long_path_prefix_is_applied(tmp_path: Path) -> None:
    service = DossierSimulationWorkspaceService()
    sample = tmp_path / ("a" * 30) / ("b" * 30) / ("c" * 30) / "archivo.pdf"

    prefixed = service._with_windows_long_path_prefix(sample)

    assert prefixed.startswith('\\\\?\\')


def test_simulation_creates_root_even_when_no_files_can_be_copied(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "simulation-empty-root.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    config = _build_config(root_path=root_path, excel_path=workbook_path)
    summary = DossierService().run_simulation(config, report_dir=tmp_path / "reports-empty-root")
    assert summary.simulation_root
    assert Path(summary.simulation_root).exists()
    assert all(not item.simulation_path for item in summary.items)


def test_simulation_treats_directory_source_paths_as_missing(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "simulation-dir-source.xlsx", ["Centro Proyecto", "Número de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    config = _build_config(root_path=root_path, excel_path=workbook_path)
    config.pdf_sources[0].source_pdf_path = tmp_path / "sources"

    summary = DossierService().run_simulation(config, report_dir=tmp_path / "reports-dir-source")

    descriptivo = next(item for item in summary.items if item.rule_name == "Descriptivo de pintura")
    assert descriptivo.status in {DossierStatus.ERROR, DossierStatus.SKIPPED}
    assert summary.simulation_root
    assert Path(summary.simulation_root).exists()

def test_folder_7_documents_are_routed_to_ensayos_not_planos(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "folder7-target.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    summary = DossierService().run_simulation(_build_config(root_path=root_path, excel_path=workbook_path), report_dir=tmp_path / "reports-folder7-target")
    folder7_items = [item for item in summary.items if item.rule_name in {"Resultados adherencia", "Informe laboratorio"}]
    assert folder7_items
    assert all("7 Ensayos" in item.target_folder for item in folder7_items)
    assert all("2 Planos" not in item.target_folder for item in folder7_items)


def test_distribution_uses_distinct_sources_for_folder_7_rules(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "folder7.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    summary = DossierService().run_simulation(_build_config(root_path=root_path, excel_path=workbook_path), report_dir=tmp_path / "reports-folder7")
    comunicado = next(item for item in summary.items if item.rule_name.startswith("Resultados"))
    informe = next(item for item in summary.items if item.rule_name.startswith("Informe"))
    assert "7.1 Comunicado" in comunicado.source_pdf_path and "Resultados.pdf" in comunicado.source_pdf_path
    assert "7.2 Informe de Ensayo Laboratorio.pdf" in informe.source_pdf_path




def test_real_distribution_skips_documents_when_rerun_on_same_folder(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "rerun-skip.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    legacy_path = cp_folder / "06_DOSSIER" / "5 Procedimiento de fabricacion" / "5.2 Procedimiento Aplicacion Pintura Tanques 2023.03.08.pdf"
    _create_pdf(legacy_path, b"%PDF-1.4\nold\n")
    config = _build_config(root_path=root_path, excel_path=workbook_path, simulation_only=False, replace_existing=True)

    first_summary = DossierService().run_distribution(
        config,
        confirm_real=True,
        report_dir=tmp_path / "reports-rerun-first",
    )
    second_summary = DossierService().run_distribution(
        config,
        confirm_real=True,
        report_dir=tmp_path / "reports-rerun-second",
    )

    assert any(item.status == DossierStatus.REPLACED for item in first_summary.items if item.rule_name == "Descriptivo de pintura")
    actionable_items = [item for item in second_summary.items if item.rule_name in {"Descriptivo de pintura", "Certificado trazabilidad", "Resultados adherencia", "Informe laboratorio"}]
    assert actionable_items
    assert all(item.status == DossierStatus.SKIPPED for item in actionable_items)
    assert all("ya está correcto" in item.observation for item in actionable_items)


def test_real_distribution_skips_equivalent_existing_folder7_document(tmp_path: Path) -> None:
    workbook_path = _write_workbook(tmp_path / "existing-folder7-equivalent.xlsx", ["Centro Proyecto", "Numero de serie"], [["CP-01", "S-001"]])
    root_path = tmp_path / "root"
    cp_folder = _build_cp_tree(root_path, "CP-01", "S-001")
    _create_all_sources(tmp_path)
    config = _build_config(root_path=root_path, excel_path=workbook_path, simulation_only=False, replace_existing=True)

    mojibake_name = f"7.1 Comunicado t{chr(0x00c3)}{chr(0x00a9)}cnico - Resultados de adherencia.pdf"
    existing_target = cp_folder / "06_DOSSIER" / "7 Ensayos" / mojibake_name
    _create_pdf(existing_target)

    summary = DossierService().run_distribution(
        config,
        confirm_real=True,
        report_dir=tmp_path / "reports-existing-folder7-equivalent",
    )

    comunicado = next(item for item in summary.items if item.rule_name == "Resultados adherencia")
    assert comunicado.status == DossierStatus.SKIPPED
    assert "ya está correcto" in comunicado.observation
    folder7_files = list((cp_folder / "06_DOSSIER" / "7 Ensayos").glob("7.1*.pdf"))
    assert len(folder7_files) == 1


def test_run_simulation_filters_by_multiple_cp_and_serie_values(tmp_path: Path) -> None:
    workbook_path = _write_workbook(
        tmp_path / "equipos-multi-filter.xlsx",
        ["Centro Proyecto", "Numero de serie"],
        [["CP-01", "S-001"], ["CP-02", "S-002"], ["CP-03", "S-003"]],
    )
    root_path = tmp_path / "orders"
    _build_cp_tree(root_path, "CP-01", "S-001")
    _build_cp_tree(root_path, "CP-02", "S-002")
    _build_cp_tree(root_path, "CP-03", "S-003")
    _create_all_sources(tmp_path)

    summary = DossierService().run_simulation(
        _build_config(root_path=root_path, excel_path=workbook_path, cp_filters=["CP-01", "CP-03"], serie_filters=["S-001", "S-003"]),
        report_dir=tmp_path / "reports-multi-filter-list",
    )

    rows = {item.row_number for item in summary.items if item.rule_name == "Descriptivo de pintura"}
    assert rows == {2, 4}
