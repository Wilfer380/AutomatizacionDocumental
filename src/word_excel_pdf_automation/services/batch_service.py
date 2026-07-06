from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

from ..config import OUTPUT_PDF_DIR, OUTPUT_REPORT_DIR, OUTPUT_WORD_DIR, PLACEHOLDER
from ..models import BatchOptions, BatchSummary, GenerationResult, SeriesRow
from ..utils.files import ensure_directory
from ..utils.text import sanitize_filename


logger = logging.getLogger(__name__)


class BatchService:
    def __init__(self, template_service, pdf_converter, report_service) -> None:
        self.template_service = template_service
        self.pdf_converter = pdf_converter
        self.report_service = report_service

    def run(
        self,
        options: BatchOptions,
        rows: list[SeriesRow],
        progress_callback=None,
    ) -> BatchSummary:
        output_dir = ensure_directory(options.output_dir)
        word_dir = ensure_directory(output_dir / OUTPUT_WORD_DIR)
        pdf_dir = ensure_directory(output_dir / OUTPUT_PDF_DIR)
        report_dir = ensure_directory(output_dir / OUTPUT_REPORT_DIR)

        results: list[GenerationResult] = []
        ordered_rows = sorted(rows, key=lambda row: row.row_number)
        total = len(ordered_rows)
        if total == 0:
            report_path = self.report_service.write_csv_report(report_dir, results)
            return BatchSummary(total=0, succeeded=0, failed=0, skipped=0, report_path=report_path, generated_items=results)

        for index, row in enumerate(ordered_rows, start=1):
            if progress_callback:
                progress_callback(index - 1, total, f"Procesando fila {row.row_number}")

            timestamp = datetime.now().isoformat(timespec="seconds")
            base_name = f"{index:04d}_{row.row_number:04d}_{sanitize_filename(row.series)}"
            docx_path = word_dir / f"{base_name}.docx"
            pdf_path = pdf_dir / f"{base_name}.pdf"
            try:
                filled = self.template_service.create_filled_copy(
                    template_path=options.template_path,
                    destination_path=docx_path,
                    replacement=row.series,
                    placeholder=options.placeholder or PLACEHOLDER,
                )
                if not filled:
                    raise RuntimeError("No se reemplazó el marcador de la plantilla.")
                generated_pdf = self.pdf_converter.convert(docx_path, pdf_dir)
                observation = row.observation if row.observation else "Correcto"
                result = GenerationResult(
                    row_number=row.row_number,
                    series=row.series,
                    docx_path=str(docx_path),
                    pdf_path=str(generated_pdf),
                    status="generated",
                    observation=observation,
                    timestamp=timestamp,
                )
                results.append(result)
            except Exception as exc:  # pragma: no cover - exercised through runtime behavior
                logger.exception("Error al procesar la fila %s (%s)", row.row_number, row.series)
                results.append(
                    GenerationResult(
                        row_number=row.row_number,
                        series=row.series,
                        docx_path=str(docx_path) if docx_path.exists() else "",
                        pdf_path=str(pdf_path) if pdf_path.exists() else "",
                        status="failed",
                        observation=str(exc),
                        timestamp=timestamp,
                    )
                )

            if progress_callback:
                progress_callback(index, total, f"Fila {row.row_number} terminada")

        report_path = self.report_service.write_csv_report(report_dir, results)
        succeeded = sum(1 for item in results if item.status == "generated")
        failed = sum(1 for item in results if item.status == "failed")
        skipped = sum(1 for item in results if item.status == "skipped")
        return BatchSummary(
            total=total,
            succeeded=succeeded,
            failed=failed,
            skipped=skipped,
            report_path=report_path,
            generated_items=results,
        )
