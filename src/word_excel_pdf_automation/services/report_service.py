from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

from ..models import GenerationResult


class ReportService:
    def write_csv_report(self, report_dir: Path, results: list[GenerationResult]) -> Path:
        report_dir.mkdir(parents=True, exist_ok=True)
        report_path = report_dir / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with report_path.open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=[
                    "row_number",
                    "series",
                    "docx_filename",
                    "pdf_filename",
                    "docx_output",
                    "pdf_output",
                    "status",
                    "observation",
                    "timestamp",
                ],
            )
            writer.writeheader()
            for item in results:
                writer.writerow(
                    {
                        "row_number": item.row_number,
                        "series": item.series,
                        "docx_filename": item.docx_filename,
                        "pdf_filename": item.pdf_filename,
                        "docx_output": item.docx_path,
                        "pdf_output": item.pdf_path,
                        "status": item.status,
                        "observation": item.observation,
                        "timestamp": item.timestamp,
                    }
                )
        return report_path
