<!-- Purpose: Significant decisions and the reasoning behind them. Kept separate so rationale does not get buried in the Progression_Log. -->

## Decision Log

### Excel/VBA as the platform
Inherited from Alex's existing build. Not re-evaluated at this stage — the two new project builds will follow the same approach for consistency and to keep scope manageable.

The fuller reasoning: the preferred approach would be a Python/Streamlit application, but Alex is on paternity leave and it would be disrespectful to rebuild the system mid-year in his absence. He will return to an already-changed landscape (AI-first working practices have accelerated since he left) and the right thing is to deliver within his paradigm. A rebuild conversation can happen when he returns.

### Exploratory phase before planning
Alex is on paternity leave and the existing codebase is not fully understood. An exploratory review of his build is the necessary first step before any planning or development begins.

### Separation of documentation and interpretation
The exploratory phase is split into two discrete sessions: one to document Alex's tool (pure recording, no interpretation), and one to interpret what has been documented (collaborative, with the user's knowledge of the codebase applied). This prevents interpretive assumptions baking in before they have been validated.

### Reference folder as holding area
Alex's .xlsm, exported VBA modules, user instructions, and the two new project templates will be held in a reference/ folder, clearly distinct from code_base/ which is reserved for the actual builds.

### User instructions treated as 95% reliable
Alex's user instructions are a valuable source of intent and user experience context, but are not treated as ground truth. The tool itself is the 100% reliable reference. Anything load-bearing from the instructions is verified against the code and template directly.

### Test database as a non-negotiable gate
No real data will touch the tool until it has been verified against the test database. Test database access and API endpoint verification is a discrete session that gates all build work.

### Credentials to be supplied via Config sheet
Credentials will not be hardcoded in new tools. A dedicated section of the `Config` sheet in the Excel workbook will hold username and password fields, clearly labelled. This avoids plaintext credentials in committed code and keeps the tool self-contained without requiring a separate credentials file.

### Live/Test toggle pattern carried forward
Alex's Live/Test toggle mechanic is a proven pattern that solves a real operational need. It will be replicated in the new tools, implemented as a data validation drop-down on the `Config` sheet (`Test` / `Live`) rather than Alex's named range on the `Orgs` sheet — consistent with the move to a dedicated Config sheet.

### read_excel tool added to MCP server
Direct inspection of `.xlsm` and `.xlsx` files requires openpyxl running on the Windows machine, not in the bash container. A `read_excel` tool was added to the file-reader MCP server to bridge this gap. openpyxl 3.1.5 installed on the Windows machine. This tool is now the standard method for all Excel file inspection in this project.

### Session 3 voided and scheduled for redo
The documentation of Alex's tool in Session 3 was based on `.bas` files only. The `.xlsm` workbook was never directly inspected despite this being the stated intent. The session is treated as incomplete and was redone in full using `read_excel` combined with the `.bas` modules.

### One generalised VBA codebase, three workbook instances
Rather than two entirely separate builds, the VBA is written once as a generalised codebase. All project-specific variation is held in the workbook (Config sheet, Drop downs sheet, Orgs sheet). Three instances are produced: Base (unpopulated template), Managing Frailty (Project 35), Virtual Ward (Project 68). This reduces maintenance burden and keeps the logic consistent across both tools.

### Config sheet as the single configuration point
Alex's tool uses named ranges scattered across sheets (ProjectID on Home, Toggle on Orgs, SubmissionFolder on Home). The new tools consolidate all configuration into a single `Config` sheet with clearly labelled rows. The VBA reads exclusively from named ranges pointing to Config cells. This makes the tool easier to configure, easier to hand to Alex, and easier to adapt in future.

### Orientation toggle in Config
The new templates have a transposed layout (questions in rows, patients in columns) compared to Alex's tool (questions in columns, patients in rows). Rather than hardcoding orientation per tool, an `Orientation` named range on Config (`Columns` / `Rows`) controls how the importer iterates the template. This future-proofs the VBA against further template variations.

### ServiceID = 0 for new projects
Both Managing Frailty (Project 35) and Virtual Ward (Project 68) use ServiceID = 0. This is held as a configurable value in the Config sheet, not hardcoded — consistent with the principle that all project-specific values live in the workbook.

### .xlsx for workbook creation, .xlsm after VBA import
openpyxl cannot produce a valid `.xlsm` file — saving with that extension produces a corrupt file. Workbooks are created as `.xlsx` via the `write_excel` MCP tool. The user opens the file in Excel, saves as `.xlsm` (which adds the correct macro container structure), then imports the `.bas` modules via the VBA editor. This is a single manual step per tool instance.

### write_excel tool added to MCP server
Workbook creation and population requires programmatic Excel file writing. A `write_excel` tool was added to the file-reader MCP server using openpyxl. Supports operations including: create_workbook, add_sheet, delete_sheet, rename_sheet, write_cell, write_range, set_named_range, add_validation, set_sheet_visibility, set_column_width, set_bold, set_font_colour, set_background_colour, save_workbook. In-place editing of existing files supported (no create_workbook required). Verified working.

### reference/ and new_questionnaires/ committed to Git
Both the `reference/` folder (Alex's `.xlsm`, exported `.bas` modules, guidance document) and the `new_questionnaires/` folder (the two new project template files) are committed to the repository. These are source materials for the build and their presence in version control provides a useful audit trail. Neither folder contains credentials or sensitive personal data — credentials were redacted from Alex's files in Session 3.

### Alex's layout, structure, approach and formatting as the default
All build decisions default to following Alex's tool. Deviations are made only where the new requirements make them unavoidable. This respects Alex's work, keeps the tools feeling like a coherent family, and reduces the interpretive burden on the builder.

### Five question types: LS, YN, N, TX, DT
The new questionnaires introduce two question types not present in Alex's tool:

**TX (free text):** Present in both Managing Frailty (2 questions) and Virtual Ward (4 questions). All are "if other please specify" type fields. TX is handled identically to `N` (numeric) — the value is passed through to the API directly as a string, with no lookup or conversion. No entry in the Drop downs sheet is required.

**DT (date):** Present in Virtual Ward only — one question, `62624` ("Referral date"). Users have been instructed to enter dates in `DD/MM/YYYY` format, but the template cell type is not enforced as text. Any entry close to a date format will have been silently converted by Excel into a date serial number (a numeric value with date formatting). The handling logic is therefore:

- **If the cell value is numeric:** treat as an Excel date serial. Validate that the date falls within the expected submission window: **1 June 2026 to 31 August 2026** (Excel serials 46,174 to 46,269 inclusive). If valid, convert to `YYYY-MM-DD 00:00:00.000` format for the API. Time component is always `00:00:00.000`.
- **If the cell value is not numeric (text, empty, or anything else):** flag to the user via the standard orange cell convention.
- **If the cell value is numeric but outside the valid date range:** also flag orange.

The API expects dates in `YYYY-MM-DD HH:mm:ss.000` format, confirmed by inspecting the date column format in SSMS (value: `1900-01-01 17:17:00.000` format). The `questionType` string for DT in the API payload is still to be confirmed — this is the only remaining open item for the DT implementation.

### Folder boundary: code_base/ vs environment/
`code_base/` holds all build artefacts — `.xlsx` workbook files and `.bas` VBA modules. `environment/` holds Git infrastructure only — `git_push.py`, `git_revert.py`, `git_log.txt`. Nothing else belongs in `environment/`. Scratch and test files do not belong in either folder and should be deleted once their purpose is served.

### test_inputs/ gitignored
The `test_inputs/` folder holds local working files used for testing — valid and deliberately broken template files. It is gitignored and not committed to the repository. These files are not build artefacts and may contain real-looking data that should not be version-controlled.

### No hardcoded row or column references in VBA
All positional references in the VBA codebase use named ranges exclusively. This decouples the code from the physical layout of the workbook — cells and ranges can be moved without touching the VBA, provided named ranges are updated accordingly.

### StartCols as the single source of truth for source template positions
The `StartCols` named range (Home row 4, K onwards) holds the row or column number in the source template for every field — including the unique reference (K4) and all questions (L4 onwards). The importer reads these values to know where to find each piece of data in the submitted file. Unique reference and questions are treated identically — no special-casing.

### Alex's case code flow left untouched
The case code creation, posting, and closing sequence in `A3_API_Calls` is carried forward from Alex's tool without modification. It is proven in production. Changes are made only where necessary (named range refs, question type additions). The flow itself is not touched until it has been proven to fail.

### Validation parked as a separate module and session
Three categories of validation are deliberately excluded from the current codebase:
1. File validation (does the submitted file match expectations)
2. Response validation (`ResponseValidator` — orange cell colouring for invalid responses)
3. Database comparison (`CaseCodeProcessed`, `QuestionResponseMatcher` — duplicate detection and green/orange cell colouring)

All three will be addressed in dedicated sessions. Keeping them separate allows the core import and post flow to be built and tested independently.

### Build phases: Home sheet population is a separate activity
Populating the Home sheet metadata rows (question numbers, type codes, source positions, QIDs, column headers) for each project instance is a build-time activity, not part of the tool's runtime functionality. This work is scoped to the tool instance build sessions.

### Colour palette established from Alex's reference tool
Alex's formatting conventions have been reverse-engineered from `Template_Processing_Tool.xlsm` and documented in `dynamic_docs/Colour_Palette.md`. Five colours identified: panel background (`D3DAEE`), input cell (`F9FAFA`), metadata row yellow (`FFFF00`), dark navy header (`1F3864`), label grey (`F2F2F2`). These are now applied to `CCR_Tool_Base.xlsx` and serve as the standard for all subsequent sheet formatting. MCP server updated with `set_font_colour` op and opacity fix for `set_background_colour`.

### Input file flow redesigned: multi-select file picker with semi-automated matching
**Original plan:** single file path on Config; user pastes path before import; tool reads one file.

**New plan:** a single button presents a multi-select `msoFileDialogFilePicker`. The user selects one or more files. For each file, the tool validates the file structure (B5), then reads the Support sheet via XLookup for org name and submission descriptor, and matches against the Orgs sheet — automatically confirming where one match exists, prompting the user to choose where multiple exist, and skipping with a message where no match is found. `SubmissionFolderPath` (formerly `SubmissionFilePath`) remembers the last-used folder so the picker opens in the right place on subsequent runs.

**Processing sequence per file:** Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)

**Impact:** `B4_Process_Folder.bas` owns the picker and matching logic. `B5_File_Validator.bas` owns structural file validation and sits between the picker and matching. `B1_Importer.bas` is pure data transfer and is called by B4 once a match is confirmed. The three modules are kept deliberately separate.

**Rationale:** Review of Alex's code revealed folder-cycling as his actual implementation. Multi-select picker is simpler and more predictable than folder-cycling via `Dir()` — the user has explicit control over which files are processed. File validation before matching prevents the matching logic from running against files that cannot be reliably read.

### Process_Folder as orchestrator, not importer
`B4_Process_Folder.bas` handles file selection and org/submission matching only. It calls B5 for validation and B1 for import. It does not contain validation or import logic itself.

### SubmissionFilePath renamed to SubmissionFolderPath
The `SubmissionFilePath` named range and Config sheet label have been renamed to `SubmissionFolderPath`. The cell holds the folder path of the last-used file picker location. `B1_Importer.bas` still references the old name and will be updated in Session D.

### B1_Importer to accept parameters, not named ranges
`FileImporter` accepts file path, submission ID, org name, and submission name as parameters passed by `B4_Process_Folder`. It does not read from Config named ranges for these values. All validation logic has been removed — file validation is B5's responsibility, and by the time B1 is called the file has already been validated and matched.

### Separation of response validation and duplicate detection into distinct modules
Response validation (orange cell colouring) and duplicate detection (green cell colouring) are separate concerns and are implemented in separate modules:

- **B6_Response_Validator:** validates response text against the Drop downs sheet locally; no API calls required; produces orange cell colouring for invalid responses
- **B7_Duplicate_Detector:** calls the API to retrieve existing case codes and responses; compares against imported rows; produces green cell colouring for likely-already-imported rows

Rationale: the two operations have different dependencies (local vs API), different outputs, and are independently testable. Separating them keeps each module focused and makes debugging straightforward.

### Managing Frailty as primary build target; Virtual Ward as amendment
The tool is built and fully tested against Managing Frailty (Project 35) before Virtual Ward (Project 68) is introduced. Virtual Ward is treated as an amendment to the proven Managing Frailty build rather than a parallel workstream. This reduces risk: Managing Frailty has no DT questions, making it the simpler test case. Virtual Ward introduces the DT question type and will be the point at which DT handling is confirmed end-to-end.

### File validation via Config-driven checks: MandatorySheets and SpotChecks
`B5_File_Validator` is config-driven rather than hardcoded. Two new named ranges on the Config sheet control what is checked:

- **`MandatorySheets`** (Config!B14): `^`-delimited list of sheet names that must be present in the submitted file (e.g. `CCR^Support`)
- **`SpotChecks`** (Config!B15): `^`-delimited list of cell checks, each in the format `SheetName!CellRef:ExpectedValue` (e.g. `CCR!BB4:Patient 50`)

This design means validation rules can be adjusted per tool instance by changing Config values only — no VBA changes required. SpotChecks are targeted at structural anchors (first/last patient column headers, section headers, question type labels) to catch row and column deletions or insertions that would break the importer.

The Support sheet check (Check 2) is the only non-configurable check — it uses XLookup to validate Project ID, Submission Period, Spec Type, and Organisation Name against Config values and known constants. This is the same pattern used by B4 to read org name and submission descriptor.

### openpyxl drops Form Controls on save — .xlsm hands-off once buttons are present
openpyxl silently drops Excel Form Controls (buttons) when saving a `.xlsm` file, even when loading an existing file rather than creating a new one. Once buttons have been added to the workbook, the `.xlsm` must never be passed through the MCP `write_excel` tool. All future workbook changes are applied directly in Excel by the user, with Claude providing an exact instruction list of values, named ranges, and formatting to apply.

### B1 empty-record skip logic: has-data check across all question cells
The original B1 skip logic checked whether the unique reference cell was blank. This was ineffective because unique references ("Patient 1", "Patient 2" etc.) are hardcoded headers in the template — they are never blank regardless of whether any response data exists.

The replacement logic iterates all question positions in `StartCols` (index 2 onwards, skipping the unique ref at index 1) and checks whether at least one response cell is non-blank. Threshold is 1: any single non-blank response constitutes a real record and must be imported. This works for both Columns and Rows orientations — the only difference is argument order in `Wsh_Source.Cells()`. An `Exit For` on first hit keeps performance acceptable even for large DataMax values.

### B6 must precede the first live API test
B6_Response_Validator must be built and verified before any records are posted to the database. The API import will fail if responses are invalid — for example, an LS response that does not match a valid list item ID cannot be converted and will cause the post to error. Orange cells are the user's signal to correct data before posting. Running a post against unvalidated data risks creating partial or corrupt case codes against the test database.

### B6 response validation scope: current-run rows only
B6 validates only the rows imported in the current run, not all rows present on the Home sheet. B4 passes the first and last row numbers written during the run (tracked via ByRef parameters on B1). This is the correct scope because: (a) previously imported rows may have been deliberately corrected by the user and should not be re-validated; (b) orange cell state from a previous run is cleared by the `.Clear` operation at the start of a new run.

### B6 validates all five question types; no assumptions about template constraints
B6 applies active validation rules to all five question types. Template drop-down constraints (e.g. Yes/No for YN questions) are not relied upon — users may have edited cells directly or submitted files with unexpected values. The rules are:

- **LS:** response text must match a valid option in the Drop downs sheet (even column = text column; odd column = list item IDs)
- **YN:** must be exactly "Yes" or "No"
- **N:** must be numeric (IsNumeric check)
- **TX:** must not be numeric; any non-blank text string is valid
- **DT:** must be a numeric Excel date serial within 46174–46269 (1 Jun–31 Aug 2026)

### Clear uses .Clear not .ClearContents; borders reapplied after clear
When the user chooses to start fresh, `FullDataArea.Clear` is used rather than `.ClearContents`. `.Clear` wipes both cell values and formatting, ensuring orange validation colouring from a previous run does not persist. After clearing, thin grid borders are reapplied to `FullDataArea` using the six explicit border constants (`xlEdgeTop`, `xlEdgeBottom`, `xlEdgeLeft`, `xlEdgeRight`, `xlInsideHorizontal`, `xlInsideVertical`). Diagonal borders (`xlDiagonalDown`, `xlDiagonalUp`) are deliberately excluded.

### Error Log sheet: minimalist export design
The Error Log sheet is intentionally plain — no panel colours, no formatting beyond the user-applied bold on row 1. It is a diagnostic export tool, not a user-facing sheet. One row per validation error: Row, Unique Ref, Question ID, Question No., Question Type, Invalid Value. The sheet is cleared and re-headed at the start of each B6 run.

### Case code column holds the case code returned by the API
Column J on the Home sheet receives the case code written back by A3 after a successful post. It does not hold a unique reference string or any other value. The unique reference ("Patient 1" etc.) is in column K. This is consistent with A3's existing logic: `Rng_Cell.Offset(0, -1).Value = Str_CaseCode`.
