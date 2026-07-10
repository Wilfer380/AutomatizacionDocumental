# Design: Phase 2 Real Dossier Distribution

## Technical Approach

Preserve the Tkinter Phase 2 panel, Phase 1 generation, and current service stack, but split Phase 2 into planning and execution. `DossierService` orchestrates Excel loading, CP/Serie validation, source-PDF planning, reporting, and calls a new `DossierDistributionService` only for confirmed real writes. Simulation and real mode use the same `DossierActionResult` plan, so the real process cannot invent different destinations.

## Architecture Decisions

| Decision | Choice | Alternatives considered | Rationale |
|---|---|---|---|
| Real write boundary | Create `services/dossier_distribution_service.py`; keep `DossierFilePlacerService` plan-only. | Put `shutil.copy2` in the placer. | A dedicated boundary keeps simulation vs real execution auditable and prevents accidental writes during planning. |
| Confirmation contract | `run_distribution(config, confirm_real=False)` raises `RealModeConfirmationRequiredError` when `config.simulation_only is False`. | Return a skipped summary. | A hard error is safer for shared folders because an unconfirmed real request is a caller bug, not a business skip. |
| Source PDFs | The four Phase 2 PDFs come from Section 3 rows in `DossierPanel` (`DocumentRow.source_var`, target folder, final name, action, mandatory). | Derive source PDFs from CP folders. | The UI already exposes explicit source-file selection; using it makes operator intent visible and testable. |
| Phase 1 routing | Optional routing consumes Phase 1 `BatchSummary.generated_items` / `GenerationResult.pdf_path` as an extra route request. | Modify Phase 1 generation to write into dossier folders. | Phase 1 remains isolated; Phase 2 owns dossier placement and cross-checks. |
| Test bootstrap | Add minimal `pytest` setup now. | Keep manual-only validation. | The repo currently has no test runner; real copy/backup logic needs repeatable temporary-filesystem tests. |

## Data Flow

```text
UI root + Excel + 4 PDF rows
 -> DossierService.run_distribution()
 -> DossierValidatorService: CP/Serie headers, unique CP, 06_DOSSIER, Planos, Serie in Planos
 -> DossierFilePlacerService: planned actions for folders 5/6/7 with source_pdf_path
 -> simulation: JSON report only
 -> real + confirm_real: DossierDistributionService backup-before-replace + copy
 -> DossierReportService: row/action/result JSON
```

`execute_real_process()` must show confirmation, build `DossierConfig(simulation_only=False)`, and call `run_distribution(..., confirm_real=True)` in the existing background worker/queue pattern.

## File Changes

| File | Action | Description |
|---|---|---|
| `src/word_excel_pdf_automation/services/dossier_distribution_service.py` | Create | Executes confirmed copy/replace from planned actions and creates destination-file backups first. |
| `src/word_excel_pdf_automation/services/dossier_service.py` | Modify | Add `run_distribution()`, `RealModeConfirmationRequiredError`, source-plan orchestration, and distributor injection. |
| `src/word_excel_pdf_automation/services/dossier_validator_service.py` | Modify | Enforce unique CP, `06_DOSSIER`, `Planos`, and Serie presence under `Planos`. |
| `src/word_excel_pdf_automation/services/dossier_file_placer_service.py` | Modify | Build actions only for `5 Pintura`, `6 Trazabilidad`, `7 Ensayos`; attach `source_pdf_path`. |
| `src/word_excel_pdf_automation/services/dossier_backup_service.py` | Modify | Add destination-file backup API before overwrite. |
| `src/word_excel_pdf_automation/services/dossier_report_service.py` | Modify | Record mode, source PDF, backup path, written path, skipped reason, Phase 1 route results. |
| `src/word_excel_pdf_automation/dossier_models.py` | Modify | Add `DossierPdfSource`, action type, execution mode, `source_pdf_path`, backup/written/skipped fields. |
| `src/word_excel_pdf_automation/ui/dossier_panel.py` | Modify | Pass the four PDF rows into config/plan, wire confirmation, keep background execution. |
| `config.json`, `config.example.json`, `README.md` | Modify | Document folder rules, confirmation, source-PDF rows, Phase 1 optional routing. |
| `requirements-dev.txt`, `pytest.ini`, `tests/conftest.py`, `tests/test_dossier_distribution.py` | Create | Minimal pytest bootstrap and filesystem/service tests. |

## Interfaces / Contracts

`DossierService.run_distribution(config: DossierConfig, confirm_real: bool, report_dir: Path | None = None, phase1_items: list[GenerationResult] | None = None) -> DossierRunSummary`.

For the four Phase 2 PDFs, each UI row becomes `DossierPdfSource(document_name, source_pdf_path, target_folder, final_name_pattern, action, mandatory)`. The plan/result model stores `source_pdf_path`, `target_folder`, `planned_path`, `action_type`, `execution_mode`, `backup_path`, `written_path`, and `skipped_reason`. Missing mandatory source files create row/action `ERROR`; missing optional sources create `SKIPPED`.

Optional Phase 1 routing is handed off after Phase 1 generation completes, by passing `BatchSummary.generated_items` into Phase 2. Cross-check data: `GenerationResult.series`, `GenerationResult.pdf_filename/pdf_path`, dossier row `serie`, dossier row `cp`, and the validator-resolved `cp_folder`. Only exact normalized Serie plus matched CP row may route to `06_DOSSIER/6 Trazabilidad`; mismatches are `SKIPPED`.

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | confirmation error, source PDF mapping, CP/header matching, Serie-in-Planos, target names. | Add `pytest` and use `tmp_path` + openpyxl workbooks. |
| Integration | simulation writes nothing; real mode writes only inside matched CP and backs up replacements. | Temporary CP tree with sample PDFs. |
| Manual smoke | UI stays responsive; real mode prompts before writes. | Run `python run_app.py` after automated tests. |

## Migration / Rollout

No data migration required. Roll out simulation first, inspect JSON reports, then enable one confirmed real run.

## Open Questions

None.
