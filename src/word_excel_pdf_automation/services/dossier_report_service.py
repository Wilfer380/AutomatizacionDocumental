from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from ..dossier_models import DossierRunSummary


class DossierReportService:
    def write_json_report(self, report_dir: Path, summary: DossierRunSummary) -> Path:
        report_dir.mkdir(parents=True, exist_ok=True)
        report_path = report_dir / f"dossier_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with report_path.open("w", encoding="utf-8") as handle:
            json.dump(summary.to_dict(), handle, ensure_ascii=False, indent=2)
        summary.report_path = report_path
        return report_path
