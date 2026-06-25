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
  - `environment/managing_frailty_dropdowns.xlsx` — 10 LS questions, columns A–T, verified correct
  - `environment/virtual_ward_dropdowns.xlsx` — 18 LS questions, columns A–AJ, verified correct
- Both sheets formatted consistently:
  - Row 1 (QIDs): dark navy background (`2E3A87`), bold
  - Row 2 (question labels): grey-blue background (`BFC9E0`), bold
  - Data rows: alternating light blue (`EEF1F8`) / white column pairs — odd-numbered questions coloured, even-numbered white
- `write_excel` tool behaviour noted: formatting operations require data and formatting to be written in a single operation chain (tool does not open existing files for modification)
