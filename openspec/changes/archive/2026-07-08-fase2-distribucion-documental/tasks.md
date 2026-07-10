# Tasks: Phase 2 Real Dossier Distribution

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 450-650 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → PR 2 → PR 3 |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: pending
400-line budget risk: High

## Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Contracts + backup boundary | PR 1 | `dossier_models.py`, `dossier_backup_service.py`, `dossier_distribution_service.py` scaffold |
| 2 | Validation + planning/execution | PR 2 | validator, placer, service orchestration, optional Phase 1 routing |
| 3 | UI/docs/tests | PR 3 | `dossier_panel.py`, `config*.json`, `README.md`, pytest bootstrap |

## Phase 1: Foundation / Contracts

- [x] 1.1 Update `src/word_excel_pdf_automation/dossier_models.py` with distribution-aware fields for source PDF, execution mode, backup/written paths, and skipped reason.
- [x] 1.2 Extend `src/word_excel_pdf_automation/services/dossier_backup_service.py` with destination-file backup support while preserving config-backup behavior.
- [x] 1.3 Create `src/word_excel_pdf_automation/services/dossier_distribution_service.py` as the confirmed real-write boundary for copy/replace.

## Phase 2: Validation and Distribution Logic

- [x] 2.1 Tighten `src/word_excel_pdf_automation/services/dossier_validator_service.py` to require unique CP, `06_DOSSIER`, `Planos`, and Serie-in-Planos checks with row-level skip/error statuses.
- [x] 2.2 Refactor `src/word_excel_pdf_automation/services/dossier_file_placer_service.py` to build plans only for folders `5/6/7` and attach `source_pdf_path` without writing files.
- [x] 2.3 Add `run_distribution(..., confirm_real)` and confirmation-error handling in `src/word_excel_pdf_automation/services/dossier_service.py`, routing simulation to planner and real mode to distributor.
- [x] 2.4 Add optional Phase 1 routing for `BatchSummary.generated_items` / `GenerationResult.pdf_path` into `CP/6 Trazabilidad` with filename, row, and CP cross-checks.

## Phase 3: UI and Configuration Wiring

- [x] 3.1 Wire the four document rows in `src/word_excel_pdf_automation/ui/dossier_panel.py` into the distribution config and real-mode confirmation flow; keep the background worker pattern.
- [x] 3.2 Update `config.json` and `config.example.json` to document editable documental root, folder rules, and optional Phase 1 routing inputs.
- [x] 3.3 Update `README.md` with simulation/real-mode behavior, backup-before-replace, and operator cautions.

## Phase 4: Tests and Verification

- [x] 4.1 Add `requirements-dev.txt`, `pytest.ini`, and `tests/conftest.py` for a minimal pytest bootstrap.
- [x] 4.2 Add `tests/test_dossier_distribution.py` covering header detection, CP/folder validation, simulation no-write behavior, backup-before-replace, and real-mode confirmation.
- [x] 4.3 Add a smoke test for optional Phase 1 routing success and mismatch cases.

## Phase 5: Corrective CP Matching Regression Fix

- [x] 5.1 Tighten CP folder matching so missing or ambiguous CP folders do not validate into the wrong dossier folder.
- [x] 5.2 Add regression coverage for missing CP and ambiguous/misleading CP folder names.
