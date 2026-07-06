from __future__ import annotations

import logging
import shutil
import subprocess
from pathlib import Path


logger = logging.getLogger(__name__)


class PdfConverter:
    def __init__(self, libreoffice_path: str | None = None) -> None:
        self.libreoffice_path = libreoffice_path or ""

    def convert(self, docx_path: Path, output_dir: Path) -> Path:
        output_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = output_dir / f"{docx_path.stem}.pdf"

        if self._try_docx2pdf(docx_path, output_dir):
            return pdf_path

        libreoffice = self._resolve_libreoffice()
        if libreoffice:
            self._convert_with_libreoffice(docx_path, output_dir, libreoffice)
            if pdf_path.exists():
                return pdf_path

        if self._convert_with_word(docx_path, pdf_path):
            return pdf_path

        raise RuntimeError(
            "No fue posible convertir DOCX a PDF. Verifique que Microsoft Word, docx2pdf o LibreOffice estén disponibles."
        )

    def _try_docx2pdf(self, docx_path: Path, output_dir: Path) -> bool:
        try:
            from docx2pdf import convert as docx2pdf_convert
        except Exception:
            return False

        try:
            docx2pdf_convert(str(docx_path), str(output_dir))
            pdf_path = output_dir / f"{docx_path.stem}.pdf"
            return pdf_path.exists()
        except Exception as exc:  # pragma: no cover - backend dependent
            logger.warning("docx2pdf conversion failed for %s: %s", docx_path, exc)
            return False

    def _resolve_libreoffice(self) -> str | None:
        if self.libreoffice_path:
            candidate = Path(self.libreoffice_path)
            if candidate.exists():
                return str(candidate)

        for command in ("soffice", "soffice.exe"):
            resolved = shutil.which(command)
            if resolved:
                return resolved

        common_paths = [
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ]
        for candidate in common_paths:
            if candidate.exists():
                return str(candidate)
        return None

    def _convert_with_libreoffice(self, docx_path: Path, output_dir: Path, soffice: str) -> None:
        command = [
            soffice,
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(docx_path),
        ]
        subprocess.run(command, check=True, capture_output=True, text=True)

    def _convert_with_word(self, docx_path: Path, pdf_path: Path) -> bool:
        try:
            import pythoncom
            import win32com.client
        except Exception:
            return False

        word_app = None
        document = None
        try:
            pythoncom.CoInitialize()
            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            document = word_app.Documents.Open(
                str(docx_path.resolve()),
                ReadOnly=True,
                AddToRecentFiles=False,
                ConfirmConversions=False,
                Visible=False,
            )
            document.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
            return pdf_path.exists()
        except Exception as exc:
            logger.warning("La conversión con Microsoft Word falló para %s: %s", docx_path, exc)
            return False
        finally:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            if word_app is not None:
                try:
                    word_app.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
