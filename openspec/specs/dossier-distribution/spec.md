# dossier-distribution Specification

## Purpose

Define the behavior for Phase 2 dossier distribution from an editable documental root, with flexible Excel-driven CP/Serie matching, safe real execution, and auditable row-level outcomes.

## Requirements

### Requirement: Editable root and flexible workbook headers

The system MUST let the user select the documental root path and MUST detect CP and Serie columns from Excel headers even when labels vary.

#### Scenario: Headers are found under alternate names
- GIVEN a workbook with CP and Serie columns named differently
- WHEN the user loads the file
- THEN the system MUST identify both fields and continue

#### Scenario: A required header is missing
- GIVEN a workbook without a detectable CP or Serie column
- WHEN the user loads the file
- THEN the system MUST mark the row set as blocked and report the missing field

### Requirement: CP folder discovery and dossier validation

The system MUST find the CP folder under the selected root, MUST validate `06_DOSSIER` and `Planos`, and MUST confirm the Serie exists somewhere inside `Planos`.

#### Scenario: The CP folder structure is valid
- GIVEN a matching CP folder under the root
- WHEN validation runs
- THEN the system MUST accept the row for distribution

#### Scenario: The series is not present in Planos
- GIVEN a CP folder with `06_DOSSIER` and `Planos` but no matching Serie under `Planos`
- WHEN validation runs
- THEN the system MUST skip the row and record the missing Serie status

### Requirement: Real distribution into CP folders

The system MUST place the four defined Phase 2 PDFs only inside the matched CP folder, using folders `5`, `6`, and `7` as defined by the dossier rules.

#### Scenario: Distribution succeeds
- GIVEN a validated CP row
- WHEN real distribution runs
- THEN the system MUST copy the PDFs into the target subfolders under that CP folder

#### Scenario: The CP match is ambiguous or missing
- GIVEN no unique CP folder match
- WHEN distribution runs
- THEN the system MUST NOT write files outside the matched CP folder and MUST record the row as skipped

### Requirement: Backup before replace

The system MUST create a backup before replacing any existing destination file.

#### Scenario: A destination file already exists
- GIVEN a target file already present
- WHEN real distribution replaces it
- THEN the system MUST preserve a backup before writing the new file

### Requirement: Simulation and real mode confirmation

The system MUST support simulation mode and real mode. Real mode MUST require explicit user confirmation before any write occurs.

#### Scenario: Simulation mode is selected
- GIVEN simulation mode
- WHEN the user executes the flow
- THEN the system MUST report planned actions without changing files

#### Scenario: Real mode is confirmed
- GIVEN real mode and user confirmation
- WHEN the user executes the flow
- THEN the system MUST perform the write actions

### Requirement: Optional Phase 1 PDF routing

The system MAY route Phase 1 generated PDFs to `CP/6 Trazabilidad` when the filename, Excel row, and found CP cross-check agree.

#### Scenario: Cross-check succeeds
- GIVEN a Phase 1 generated PDF and matching Excel row
- WHEN routing runs
- THEN the system MUST allow placement in `CP/6 Trazabilidad`

#### Scenario: Cross-check fails
- GIVEN a filename or CP mismatch
- WHEN routing runs
- THEN the system MUST skip the file and record the mismatch

### Requirement: Row-level reporting and continuation

The system MUST report status per row, including planned actions, skips, backups, and write results, and MUST continue with later rows when a row is missing CP, dossier, Serie, or destination folder.

#### Scenario: Multiple rows contain errors
- GIVEN a batch with valid and invalid rows
- WHEN processing runs
- THEN the system MUST continue after each invalid row and summarize each row status
