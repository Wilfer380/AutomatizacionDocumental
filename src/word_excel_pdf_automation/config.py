from __future__ import annotations

import json
from pathlib import Path
from typing import Any


APP_NAME = "AutomatizaciónDocumental"
PLACEHOLDER = "[Serie]"
OUTPUT_WORD_DIR = "Word_generados"
OUTPUT_PDF_DIR = "PDF_generados"
OUTPUT_REPORT_DIR = "Reportes"
DEFAULT_CONFLICT_STRATEGY = "overwrite"
LOG_DIR = "logs"
LOG_FILE = "automatizaciondocumental.log"

PROJECT_ROOT = Path(__file__).resolve().parents[2]

DEFAULT_TEMPLATE_PATH = Path(r"C:\Users\wandica\Downloads\Trazabilidad descrptivo pintura - serie.docx")
DEFAULT_EXCEL_PATH = Path(r"C:\Users\wandica\Downloads\Listado equipos SLA COL.xlsx")
DEFAULT_OUTPUT_DIR = Path(r"C:\Users\wandica\Downloads\word_excel_pdf_output")
DEFAULT_DOSSIER_SIMULATION_DIR = DEFAULT_OUTPUT_DIR / "DossierSimulation"

DEFAULT_DOSSIER_ROOT_PATH = Path(r"Q:\GROUPS\CO_MDE_DISENO_DI\01_ORDERS\02_DOCUMENTS_APPROVAL_CERTIFIED")
DEFAULT_DOSSIER_EXCEL_PATH = Path(r"C:\Users\wandica\Downloads\Listado equipos SLA COL.xlsx")
DEFAULT_DOSSIER_CONFIG_PATH = PROJECT_ROOT / "config.json"
DEFAULT_DOSSIER_CONFIG_EXAMPLE_PATH = PROJECT_ROOT / "config.example.json"

DOSSIER_CP_SYNONYMS = (
    "cp",
    "c.p.",
    "centro proyecto",
    "centro de proyecto",
    "codigo cp",
    "código cp",
    "proyecto cp",
)

DOSSIER_SERIE_SYNONYMS = (
    "serie",
    "serial",
    "numero de serie",
    "número de serie",
    "no serie",
    "nro serie",
    "equipo serie",
)

SERIAL_SYNONYMS = (
    "serie",
    "serial",
    "numero de serie",
    "número de serie",
    "no serie",
    "nro serie",
    "equipment serial",
    "equip serial",
)


def read_json_file(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return data if isinstance(data, dict) else {}


def write_json_file(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
