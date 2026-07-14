from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from ..dossier_models import DossierRule, DossierRow
from ..utils.text import sanitize_filename, sanitize_windows_filename

_FOLDER6_PATTERN = re.compile(r"^\s*(?P<major>\d+)\.(?P<section>\d+)\.(?P<item>\d+)\s+(?P<title>.+?)\.pdf\s*$", re.IGNORECASE)
_NUMERIC_PREFIX_PATTERN = re.compile(r"^\s*\d+(?:\.\d+)+\s+")
_SERIE_PATTERN = re.compile(r"(?<!\d)(?P<serie>\d{6,})(?!\d)")


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

    def build_phase1_folder6_filename(
        self,
        destination_folder: Path,
        source_filename: str,
        serie: str,
        *,
        section_number: int | None = None,
        reserved_item_numbers: set[int] | None = None,
    ) -> str:
        section = section_number or self.next_folder6_section(destination_folder)
        item_number = self.resolve_folder6_item_number(destination_folder, serie, reserved_item_numbers=reserved_item_numbers)
        base_name = self.strip_numeric_prefix(source_filename)
        source_path = Path(base_name)
        stem = sanitize_windows_filename(source_path.stem, fallback=f"Certificado de Trazabilidad SN {serie}")
        suffix = source_path.suffix or ".pdf"
        return f"6.{section}.{item_number} {stem}{suffix}"

    def next_folder6_section(self, destination_folder: Path) -> int:
        sections: set[int] = set()
        if destination_folder.exists():
            for candidate in destination_folder.glob("*.pdf"):
                match = _FOLDER6_PATTERN.match(candidate.name)
                if not match:
                    continue
                if int(match.group("major")) != 6:
                    continue
                sections.add(int(match.group("section")))
        return (max(sections) + 1) if sections else 1

    def resolve_folder6_item_number(self, destination_folder: Path, serie: str, *, reserved_item_numbers: set[int] | None = None) -> int:
        reserved = set(reserved_item_numbers or set())
        if destination_folder.exists():
            per_series = Counter()
            used_items: set[int] = set(reserved)
            for candidate in destination_folder.glob("*.pdf"):
                match = _FOLDER6_PATTERN.match(candidate.name)
                if not match or int(match.group("major")) != 6:
                    continue
                item_number = int(match.group("item"))
                used_items.add(item_number)
                if serie and serie in candidate.name:
                    per_series[item_number] += 1
            if per_series:
                return sorted(per_series.items(), key=lambda entry: (-entry[1], entry[0]))[0][0]
            next_item = 1
            while next_item in used_items:
                next_item += 1
            return next_item
        return 1 if 1 not in reserved else max(reserved) + 1

    @staticmethod
    def strip_numeric_prefix(filename: str) -> str:
        source = Path(filename or "document.pdf")
        clean_name = _NUMERIC_PREFIX_PATTERN.sub("", source.name).strip()
        return clean_name or source.name or "document.pdf"
