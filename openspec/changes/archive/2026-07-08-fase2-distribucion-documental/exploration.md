## Exploration: fase2-distribucion-documental

### Current State
Phase 2 is already separated in the UI, but it is still a simulation-first dossier planner rather than a real distribution flow. `DossierPanel` seeds example rows and hardcoded previews, `DossierService.run_simulation()` only validates Excel + folder paths, writes a JSON report, and plans a config backup; `execute_real_process()` currently just shows an info message. The validator already does fuzzy CP and `06_DOSSIER` discovery, and the placer only plans target paths for configured rules like `2 Planos`, `3 Memorias`, `4 Anexos`, and `5 Fotografías`.

### Affected Areas
- `src/word_excel_pdf_automation/ui/dossier_panel.py` — current Phase 2 UI, button wiring, simulation-only behavior, seeded example rows, and summary refresh.
- `src/word_excel_pdf_automation/services/dossier_service.py` — orchestration entry point; today it only loads config and runs validation/simulation/reporting.
- `src/word_excel_pdf_automation/services/dossier_validator_service.py` — already handles flexible CP lookup, workbook parsing, and `06_DOSSIER` discovery.
- `src/word_excel_pdf_automation/services/dossier_file_placer_service.py` — currently only computes planned paths; no file copy/replace/backups execution.
- `src/word_excel_pdf_automation/services/dossier_sequence_service.py` — filename/path derivation for dossier items.
- `src/word_excel_pdf_automation/services/dossier_backup_service.py` — only backs up the JSON config, not dossier content.
- `src/word_excel_pdf_automation/services/dossier_report_service.py` — JSON report only; no operational trace of real copy actions yet.
- `src/word_excel_pdf_automation/dossier_models.py` — current dossier schema is generic and still oriented to planned actions.
- `config.json`, `config.example.json`, `README.md` — current documented defaults still describe the old simulation-oriented dossier foundation.

### Approaches
1. **Refactor the existing dossier stack into a real distributor** — keep the current UI shell, but replace the simulation-only service path with actual CP/`06_DOSSIER`/Planos/Trazabilidad routing, copy/replace logic, backups, and actionable reports/logs.
   - Pros: smallest architectural disruption; reuses current validation and fuzzy search.
   - Cons: requires careful surgery in existing services and UI flow.
   - Effort: Medium

2. **Introduce a dedicated distribution service and narrow the UI to it** — keep validation helpers, but add a separate flow object that owns source discovery, real copy/replacement, confirmations, and optional Phase 1 PDF auto-routing.
   - Pros: cleaner separation between validation and execution; easier to evolve.
   - Cons: more refactoring upfront.
   - Effort: Medium/High

### Recommendation
Take Approach 2 if the goal is to simplify Phase 2 into a stable operational workflow. The current code already has reusable validation primitives, but the execution semantics are still too loose and UI-driven. A dedicated distribution service gives a cleaner path for the real dossier flow, especially if Phase 1 PDF auto-distribution into `CP/6 Trazabilidad` becomes part of the same release.

### Risks
- Current config and model names still reflect the old broad dossier planner, so the change may need schema cleanup.
- Real copy/replace behavior can overwrite files in shared network folders; confirmation and backup rules must be explicit.
- Flexible folder matching can pick the wrong CP or dossier target if naming is inconsistent.
- Phase 1 integration could couple two workflows that are currently independent.

### Ready for Proposal
Yes. The codebase evidence is sufficient to move to proposal/spec now, with one clarification to lock down: which exact destination semantics should govern the `5`, `6`, and `7` folders and how Phase 1 PDF auto-distribution should be triggered.
