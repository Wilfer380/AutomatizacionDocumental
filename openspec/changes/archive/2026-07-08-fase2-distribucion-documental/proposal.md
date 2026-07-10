# Proposal: Phase 2 Real Dossier Distribution

## Intent
Replace the current simulation-first Phase 2 flow with a controlled real distribution process for dossier documents. The change must work from an editable documental root path, validate Excel-driven CP/Serie input, and safely place PDFs into the correct CP folder without hardcoding paths.

## Scope

### In Scope
- Editable documental root path + Excel picker with flexible CP/Serie headers.
- Flexible CP folder matching under the root; validate `06_DOSSIER` and `Planos`; verify the series exists in `Planos`.
- Real distribution of the four defined PDFs into folders `5/6/7` only inside the matched CP folder, with backup-before-replace and simulation-first execution.
- Optional routing of Phase 1 generated PDFs into `CP/6 Trazabilidad`, cross-checking filename, Excel row, and found CP.

### Out of Scope
- New dossier business rules beyond the stated Phase 2 flow.
- UI redesign, task breakdown, or application code changes in this proposal phase.

## Capabilities

### New Capabilities
- `dossier-distribution`: real Phase 2 CP/Serie validation, folder discovery, safe copy/replace, simulation, and optional Phase 1 PDF routing.

### Modified Capabilities
- `None`

## Approach
Keep the current Phase 2 UI shell, but shift execution into a real distributor path that reuses flexible validation, adds explicit confirmation before real execution, creates backups before replacement, and continues row-by-row when a CP, dossier folder, or series match is missing.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/word_excel_pdf_automation/ui/dossier_panel.py` | Modified | Phase 2 controls, simulation/real execution entry points, confirmation flow |
| `src/word_excel_pdf_automation/services/dossier_service.py` | Modified | Orchestration for validation, simulation, real distribution, and continuation on errors |
| `src/word_excel_pdf_automation/services/dossier_validator_service.py` | Modified | Flexible CP/Serie header and folder matching, `06_DOSSIER`/`Planos` checks |
| `src/word_excel_pdf_automation/services/dossier_file_placer_service.py` | Modified | Real copy/replace targeting for folders `5/6/7` and `CP/6 Trazabilidad` |
| `src/word_excel_pdf_automation/services/dossier_backup_service.py` | Modified | Backup before overwrite of destination dossier content |
| `src/word_excel_pdf_automation/services/dossier_report_service.py` | Modified | Operational trace for simulation, skips, backups, and real actions |
| `config.json`, `config.example.json`, `README.md` | Modified | Document editable paths, confirmation, and Phase 1 optional routing |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Wrong CP/folder match from flexible naming | Medium | Require CP confirmation in simulation and log skips per row |
| Overwrite on shared folders | High | Backup before replacement and explicit confirmation before real execution |
| Phase 1 coupling increases scope | Medium | Keep Phase 1 routing optional and filename/row/CP cross-checked |

## Rollback Plan
Revert to the simulation-only path, disable real copy/replace execution, and restore any replaced files from the pre-write backup copies. Keep the validation and reporting trail intact for audit.

## Dependencies
- Existing workbook parsing, CP lookup, and folder discovery helpers.

## Success Criteria
- [ ] Simulation still validates input and reports planned actions.
- [ ] Real execution writes only to the matched CP folder and continues after missing CP/dossier/series rows.
- [ ] Backups exist before any replacement.
- [ ] Optional Phase 1 routing can place PDFs into `CP/6 Trazabilidad` with cross-checks.
