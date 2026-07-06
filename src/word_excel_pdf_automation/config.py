from pathlib import Path


APP_NAME = "AutomatizaciónDocumental"
PLACEHOLDER = "[Serie]"
OUTPUT_WORD_DIR = "Word_generados"
OUTPUT_PDF_DIR = "PDF_generados"
OUTPUT_REPORT_DIR = "Reportes"
LOG_DIR = "logs"
LOG_FILE = "automatizaciondocumental.log"

DEFAULT_TEMPLATE_PATH = Path(r"C:\Users\wandica\Downloads\Trazabilidad descrptivo pintura - serie.docx")
DEFAULT_EXCEL_PATH = Path(r"C:\Users\wandica\Downloads\Listado equipos SLA COL.xlsx")
DEFAULT_OUTPUT_DIR = Path(r"C:\Users\wandica\Downloads\word_excel_pdf_output")

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
