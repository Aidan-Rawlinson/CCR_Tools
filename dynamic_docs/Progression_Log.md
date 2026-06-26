<!-- Purpose: A session-by-session history of what was built and what was decided. The project record. Authored by Claude at session end. -->

## Session 1 — 24 June 2026

**Outcome:** Project setup and briefing only. No code written.

- Git repository initialised on the CCR_Tools project folder and initial commit pushed to GitHub
- High-level brief established: the tool processes bulk uploads of Clinical Case Note Review data from Excel templates into TBN's online data collection system
- Alex's existing Excel/VBA build (for a different project) identified as the reference point
- Two new projects identified as the target scope
- Agreed approach: start with an exploratory phase to understand Alex's build before planning the adapted versions

## Session 2 — 24 June 2026

**Outcome:** Planning phase completed. No code written.

- Established working approach and division of labour: user is a competent VBA programmer, primary risk is data flow and user experience understanding rather than coding
- Confirmed API-to-database architecture: Alex's tool sends data via API to TBN's master database(s)
- Confirmed test database requirement as a non-negotiable gate before any real data is used
- Agreed to stay in Excel/VBA and follow Alex's approach faithfully — rationale documented in Decisions
- Established that Alex's user instructions exist and should be read first in the exploratory phase, weighted at ~95% reliable with the tool itself as ground truth
- Confirmed capabilities: Claude can read .xlsm structure via openpyxl (available), read exported .bas modules as plain text, and read PDFs directly — full sight of the tool is achievable
- Agreed reference folder structure for Alex's files and new project templates
- Agreed to separate documentation and interpretation into discrete sessions
- Produced a 9-session plan (see Next_Session)

## Session 3 — 25 June 2026

**Outcome:** Partial documentation only. Session voided — to be redone.

- Reference folder populated by user: `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, guidance PDF (different project clone)
- All five VBA modules read and documented
- `Alex_Tool_Reference.md` produced — but based on `.bas` files only; the `.xlsm` workbook was never directly inspected
- **Failure:** Claude stated a capability (reading `.xlsm` structure via openpyxl) in Session 2 that was not delivered in Session 3. The workbook was never passed through openpyxl. Structural knowledge of the workbook was inferred from VBA code references rather than direct inspection. This was not flagged at the time.
- The `Alex_Tool_Reference.md` is therefore incomplete and cannot be treated as a reliable foundation for build decisions
- Guidance document reviewed and discrepancies recorded — this element of the session remains valid
- Credentials redacted from `.bas` and `.xlsm` — this element of the session remains valid

## Session 4 — 25 June 2026

**Outcome:** Capability gap identified and remediated. Session plan rolled back.

- Identified that `read_excel` MCP tool was needed to properly inspect `.xlsm` and `.xlsx` files
- Built and deployed `read_excel` tool to the MCP server (openpyxl installed on Windows machine, server updated)
- Tested `read_excel` against `Template_Processing_Tool.xlsm` — confirmed working, full structural data returned
- First full inspection of `.xlsm` revealed several details absent from `Alex_Tool_Reference.md`: `Lists` sheet, full `Orgs` sheet org data, `Home` row 2 structure (question numbers), `Org_Id` broken named range
- Two new project templates identified (Managing Frailty, Virtual Ward) — inspection deferred pending session reset
- Two architectural scope changes identified for the new builds:
  - Transposed template layout (questions in rows, patients in columns) vs Alex's layout (questions in columns, patients in rows)
  - Four response types to support (`response_list_id`, `response_txt`, `response_yn`, `response_num`) vs Alex's three (`LS`, `YN`, `N`) — `TX` (free text) type to be confirmed with API team
- Decision taken: roll back to redo Session 3 properly before proceeding

## Session 3 (redo) — 25 June 2026

**Outcome:** Alex's tool fully documented. `Alex_Tool_Reference.md` rewritten and complete.

- Both `Template_Processing_Tool.xlsm` and `User_Template.xlsx` inspected in full using `read_excel` for the first time
- All five `.bas` modules already documented from Session 3 (original) — not re-read
- Alex's NACEL Submission Importer guidance PDF reviewed (now in Project Files as `GuidanceDifferent_tool.pdf`)
- `Alex_Tool_Reference.md` rewritten to incorporate all findings
- Notable findings and clarifications:
  - Question number labelling error in Home row 2 (Q51 unlabelled) — confirmed as a data entry error in the workbook
  - BM column holds a database heading stored as a question — no response data expected; blank-skip logic should handle it
  - Template has fixed capacity of 81 rows (ref numbers 101–181)
  - Tool's `Orgs` sheet is a 29-org managed subset; template's `Org list` has 95 orgs (full programme)
  - Two `Drop downs` sheets serve different purposes: template's drives Excel validation (text only); tool's drives import lookup (text + item IDs)

## Session 5 — 25 June 2026

**Outcome:** Interpretation complete. Solution designed. Build plan agreed. `write_excel` MCP tool built and verified.

- Both new questionnaire templates inspected in full via `read_excel`:
  - Managing Frailty: Project ID 35, sheet `CCR`, up to 50 patients, 42 questions, 2 narrative questions
  - Virtual Ward: Project ID 68, sheet `CCR`, up to 75 patients, 36 questions, 6 narrative questions
- Transposed orientation confirmed: questions in rows (col B), patients in columns (col E+), question type in col D
- Four question type strings confirmed from col D: `Yes/No`, `Numerical`, `Drop-down list: ...`, `Narrative`
- Section header rows confirmed (no value in col D) — must be filtered by importer
- `Support` sheet confirmed as clean machine-readable source for Project ID and submission period
- QIDs and list item IDs confirmed as absent from templates — to be supplied via SSMS CSVs
- Unique reference mechanism confirmed: "Patient 1", "Patient 2" etc. maps directly to Alex's numeric ref pattern
- Solution architecture agreed:
  - One generalised VBA codebase, no project-specific logic in code
  - Three workbook instances: Base, Managing Frailty, Virtual Ward
  - New `Config` sheet holds all configuration (replaces Alex's scattered named ranges)
  - Orientation toggle (`Columns` / `Rows`) and Environment toggle (`Test` / `Live`) as data validation drop-downs
  - ServiceID = 0 for both new projects, held in Config
  - `.xlsx` format for workbook creation (openpyxl limitation); user saves as `.xlsm` before importing `.bas` files
- Build plan agreed: 7 sessions (Sessions 5–11), gated on test database access before Session 7
- `write_excel` MCP tool designed, written, and deployed to live `server.py`
- Verified working via test file (`test_greeting.xlsx`) — all 13 operations confirmed functional
- Static spec documents populated: `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`
- `Decisions.md` updated with all new decisions from this session

## Session 6 — 25 June 2026

**Outcome:** SSMS data retrieved and saved as CSVs. QIDs cross-referenced against both templates. Drop downs build ready to proceed.

- QID data supplied for both projects and cross-referenced against questionnaire templates:
  - Managing Frailty (Project 35): 42 questions — all matched; section header rows confirmed as expected skips; encoding artefacts (apostrophes) cleaned
  - Virtual Ward (Project 68): 33 questions — all matched; section header rows confirmed as expected skips
- Question IDs cleaned: `x` prefix and `-1` suffix stripped; stored as plain integers
- List item ID data retrieved via SSMS query filtered to question IDs for both projects; cross-referenced against template drop-down values
- `colour_reference` column dropped — blank throughout, no use in build
- Data saved as four CSVs in `new_questionnaires/`:
  - `managing_frailty_question_ids.csv`
  - `managing_frailty_list_item_ids.csv`
  - `virtual_ward_question_ids.csv`
  - `virtual_ward_list_item_ids.csv`
- Each project treated as fully independent — no cross-referencing of item IDs between projects
- ⚠️ **Major flag identified:** Virtual Ward question `x62624-1` ("Referral date") is type `DT` — a date type not present in Alex's tool or Managing Frailty. Requires a new question type branch in the importer and orchestration modules, plus a date conversion function. API date format not yet confirmed. This is a blocking item for Session 8. Logged in Decisions.md and Next_Session.md.
- Minor flag: `163666` (re-admission) — database values use capital W ("Within") vs template lowercase ("within"). Database is authoritative; Drop downs sheet will use database values. Users will correct orange cells per standard workflow.

## Session 7 — 25 June 2026

**Outcome:** Drop downs sheets built and formatted for both tool instances. All DT/TX question type design decisions resolved.

- Pre-session design work:
  - TX (free text) confirmed: pass value through directly, same as N (numeric). No Drop downs entry needed.
  - DT (date) fully designed: input is Excel date serial number; validate range 1 June–31 August 2026; convert to `YYYY-MM-DD 00:00:00.000`; non-numeric or out-of-range → orange cell. API format confirmed by inspecting SSMS date column (`1900-01-01 17:17:00.000` format). API `questionType` string for DT still to be confirmed.
  - All blocking items on DT cleared except API `questionType` string.
- Drop downs sheets built from SSMS CSVs:
  - `managing_frailty_dropdowns.xlsx` — 10 LS questions, columns A–T, verified correct
  - `virtual_ward_dropdowns.xlsx` — 18 LS questions, columns A–AJ, verified correct
- Both sheets formatted consistently:
  - Row 1 (QIDs): dark navy background (`2E3A87`), bold
  - Row 2 (question labels): grey-blue background (`BFC9E0`), bold
  - Data rows: alternating light blue (`EEF1F8`) / white column pairs — odd-numbered questions coloured, even-numbered white
- `write_excel` tool behaviour noted: formatting operations require data and formatting to be written in a single operation chain (tool does not open existing files for modification)

## Session 8 — 25 June 2026

**Outcome:** Base workbook built. Folder structure cleaned up. `code_base/` established as the home for all build artefacts.

- `CCR_Tool_Base.xlsx` built in `code_base/` using `write_excel` MCP tool:
  - Five sheets: Home, Config, Orgs, Drop downs, Lists (hidden)
  - Home: column headers in row 5 (F–J), column widths set
  - Config: all 9 rows populated with labels and defaults; data validation drop-downs on Toggle (Test/Live) and Orientation (Columns/Rows)
  - All 9 named ranges defined and verified: ProjectID, ServiceID, SubmissionYear, DataSheetName, Toggle, Orientation, SubmissionFilePath, APIUsername, APIPassword
  - QuestionCols and DropDownQs deliberately excluded — instance-specific, to be defined per tool in Sessions 12–13
  - Orgs: header row only; org data instance-specific
  - Drop downs: structure note in A1; data to be copied from instance-specific sheets at build time
  - Lists: Toggle and Orientation source lists in place
- Folder structure rationalised:
  - `environment/` — confirmed as Git infrastructure only; two scratch files (`format_test.xlsx`, `test_sheet.xlsx`) deleted
  - `managing_frailty_dropdowns.xlsx` and `virtual_ward_dropdowns.xlsx` moved from `environment/` to `code_base/`
  - `code_base/` now contains all three build artefacts
- `write_excel` tool behaviour noted: `create_workbook` must be the first operation in every call (tool does not load existing files); default sheet is named `Sheet1` not `Sheet`

## Session 9 — 25 June 2026

**Outcome:** All five VBA modules written. `CCR_Tool_Base.xlsx` extended with new Config rows and Home metadata structure.

- Test database access gate resolved — toggle mechanism confirmed to route all API calls via `Toggle` named range; `Test` environment confirmed as default
- All five `.bas` modules written to `code_base/`:
  - `A1_API_SUPPORT.bas` — VBA-JSON library and UTC utilities carried forward from Alex unchanged; `GetToken()` updated to read credentials from `APIUsername`/`APIPassword` named ranges on Config
  - `A2_API_FUNCTIONS.bas` — all six API functions carried forward; `Toggle` sheet reference updated from `Orgs` to Config named range; hardcoded year replaced with `SubmissionYear` named range; `GetToken()` reference stays in A1
  - `A3_API_Calls.bas` — `PostSurveyData` loop updated: `FullDataArea` named range replaces hardcoded row/column references; `TypeCols` named range replaces `Offset(-1,0)` type code lookup; QID read directly from `QuestionCols` cell value; `TX` and `DT` question type cases added; `ServiceID`/`ProjectID` read from Config named ranges; `YN` blank check simplified
  - `B1_Importer.bas` — full rewrite of `FileImporter`; orientation-aware (`Columns`/`Rows`); reads all source positions from `StartCols` named range; `FullDataArea` cleared before import; file existence and sheet name validated before opening; `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` deliberately excluded — parked for validation module session
  - `B2_Toggle.bas` — `Toggle` sheet reference updated to Config named range; dead code removed
- `CCR_Tool_Base.xlsx` extended (by user in Excel, verified by Claude):
  - Config rows 12–13 added: `DataStart` and `DataMax` with named ranges and guidance notes in column C
  - Home metadata structure extended: row 4 (`StartCols`) added for source row/column positions; J4 populated with unique ref position (5); K4–X4 populated with placeholder question positions; `TypeCols`, `StartCols`, `QuestionCols` named ranges cover rows 3–5 from J:X; `DataArea` (J7:X19408) and `FullDataArea` (F7:X19408) named ranges cover the data table
- Approach decisions made this session:
  - No hardcoded row/column references anywhere in VBA — all positions via named ranges
  - `StartCols` is the single source of truth for source template positions — unique ref and all questions treated identically
  - Alex's case code flow left entirely untouched
  - Validation (file validation, response validation, `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator`) parked as a separate module and session
  - Build phases clarified: Home sheet population (tool instance sessions) is a separate build activity, not part of the tool's runtime functionality

## Session 10 — 26 June 2026

**Outcome:** MCP server improved. Colour palette documented. Base workbook formatted. Input file flow redesigned.

- MCP server (`server.py`) updated with three improvements:
  - `read_excel`: now reads cell background colour (`background_colour` field, ARGB hex); now iterates all cells within the used range (capped at 100 rows) rather than only value-bearing cells — enabling colour reading on empty styled cells
  - `write_excel`: now supports in-place editing of existing files (loads workbook if file exists and first op is not `create_workbook`); `set_background_colour` now prepends `FF` alpha for full opacity when 6-character RGB supplied; new `set_font_colour` operation added; `set_bold` updated to preserve existing font properties
- Colour palette reverse-engineered from `Template_Processing_Tool.xlsm`:
  - Five colours identified and documented in `dynamic_docs/Colour_Palette.md`
  - Panel background `D3DAEE`, input cell `F9FAFA`, metadata yellow `FFFF00`, dark navy header `1F3864`, label grey `F2F2F2`
  - Formatting rules derived: panel fills a contiguous block; input cells lift slightly off the panel; navy headers carry white bold text; metadata rows are yellow across full width; data rows carry no fill
- `CCR_Tool_Base.xlsx` formatted to palette:
  - Home: panel `B2:E15` in `D3DAEE`; header row `F6:X6` in `1F3864` with white bold text
  - Config: header `A1:B1` in `1F3864` with white bold text; value cells `B2:B13` in `F9FAFA` (label cells `A2:A13` already `F2F2F2`)
  - Orgs: header `A1:D1` in `1F3864` with white bold text
- **Flow redesign decision:** the single-file-path import approach is being replaced. Review of Alex's code confirmed his tool cycles over a folder of files; his user guidance (our original reference) describes a different flow. We are aligning to his code, not his guidance. The new flow cycles over all files in a folder and supports the user in matching each file to the correct org and submission — automatically where confident, with user confirmation where not.
- Session plan updated: three new sessions (A, B, C) replace the previously planned Session 10 end-to-end test:
  - Session A: Write and test VBA to populate Orgs sheet from API
  - Session B: Write and test folder-cycling importer with matching support
  - Session C: Pause and update all static spec documents to reflect new flow
- `B1_Importer.bas` flagged as requiring rework for new flow — treat as starting point
- Static spec documents (`Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`) flagged as describing old flow — not to be relied upon until Session C
