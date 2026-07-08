from __future__ import annotations

from pathlib import Path

from ..dossier_models import DossierRule, DossierRow
from ..utils.text import sanitize_filename


class DossierSequenceService:
    def sort_rules(self, rules: list[DossierRule]) -> list[DossierRule]:
        return sorted((rule for rule in rules if rule.enabled), key=lambda rule: (rule.order, rule.name.lower()))

    def build_destination_filename(self, row: DossierRow, rule: DossierRule) -> str:
        context = {
            "cp": sanitize_filename(row.cp, fallback="cp"),
            "serie": sanitize_filename(row.serie, fallback="serie"),
            "row": f"{row.row_number:04d}",
            "rule": sanitize_filename(rule.name, fallback="rule"),
        }
        candidate = rule.pdf_name_pattern.format(**context)
        candidate = candidate.strip() or f"{context['serie']}.pdf"
        if not candidate.lower().endswith(".pdf"):
            candidate += ".pdf"
        return candidate

    def resolve_target_path(self, dossier_folder: Path, rule: DossierRule, row: DossierRow) -> Path:
        return dossier_folder / rule.target_folder / self.build_destination_filename(row, rule)
