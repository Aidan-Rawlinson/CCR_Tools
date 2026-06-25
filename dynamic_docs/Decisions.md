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
Workbook creation and population requires programmatic Excel file writing. A `write_excel` tool was added to the file-reader MCP server using openpyxl. Supports 13 operations: create_workbook, add_sheet, delete_sheet, rename_sheet, write_cell, write_range, set_named_range, add_validation, set_sheet_visibility, set_column_width, set_bold, set_background_colour, save_workbook. Verified working in Session 5.

### reference/ and new_questionnaires/ committed to Git
Both the `reference/` folder (Alex's `.xlsm`, exported `.bas` modules, guidance document) and the `new_questionnaires/` folder (the two new project template files) are committed to the repository. These are source materials for the build and their presence in version control provides a useful audit trail. Neither folder contains credentials or sensitive personal data — credentials were redacted from Alex's files in Session 3.

### Alex's layout, structure, approach and formatting as the default
All build decisions default to following Alex's tool. Deviations are made only where the new requirements make them unavoidable. This respects Alex's work, keeps the tools feeling like a coherent family, and reduces the interpretive burden on the builder.

### ⚠️ DT (date) question type identified — requires new handling (Virtual Ward, x62624-1)
The Virtual Ward questionnaire (Project 68) contains one question of type `DT`: "Referral date" (question ID `x62624-1`). This question type does not exist in Alex's tool or in the Managing Frailty questionnaire. It represents a date field, stored as a `DD/MM/YYYY` string in the Excel template (column D of the CCR sheet shows "Narrative", meaning free text entry by the clinician).

This introduces two requirements not present elsewhere in the build:
1. **A new question type branch** in the importer and API post logic — `DT` must be handled alongside `LS`, `YN`, `N`, and `TX`.
2. **A date conversion function** — the value entered in the template will be a date string (format `DD/MM/YYYY`, per the template guidance). Before posting to the API, this must be converted to whatever format the API expects. The correct API date format is **not yet confirmed** and must be established with the API team before Session 8.

This is a **MAJOR FLAG**. The `DT` type is unique to Virtual Ward in the current scope and was not anticipated in the original architecture. The conversion rule and API format must be agreed before the importer (`B1_Importer.bas`) and orchestration (`A3_API_Calls.bas`) modules are written. Do not begin Session 8 without this resolved.
