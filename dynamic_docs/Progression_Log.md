<!-- Purpose: Session-by-session history of what was built and what was decided. The project record. Authored by Claude at session end. -->

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
  - `environment/` — confirmed as Git infrastructure only; two scratch files deleted
  - `managing_frailty_dropdowns.xlsx` and `virtual_ward_dropdowns.xlsx` moved to `code_base/`
- `write_excel` tool behaviour noted: `create_workbook` must be the first operation in every call

## Session 9 — 25 June 2026

**Outcome:** All five VBA modules written. `CCR_Tool_Base.xlsx` extended with new Config rows and Home metadata structure.

- All five `.bas` modules written to `code_base/`
- `CCR_Tool_Base.xlsx` extended with Config rows 12–13 (`DataStart`, `DataMax`) and Home metadata row 4 (`StartCols`)
- Named ranges: `TypeCols`, `StartCols`, `QuestionCols`, `DataArea`, `FullDataArea`
- Key decisions: no hardcoded row/column references; `StartCols` as single source of truth; Alex's case code flow untouched; validation parked

## Session 10 — 26 June 2026

**Outcome:** MCP server improved. Colour palette documented. Base workbook formatted. Input file flow redesigned.

- MCP server updated: colour reading in `read_excel`, in-place editing in `write_excel`, `set_font_colour` added, opacity fix
- Colour palette documented in `Colour_Palette.md`; `CCR_Tool_Base.xlsx` formatted to palette
- **Flow redesign:** single-file-path approach replaced with folder-cycling and semi-automated org/submission matching; aligns to Alex's actual code rather than his guidance document
- Session plan updated: Sessions A, B, C replace previously planned Session 10 end-to-end test
- `B1_Importer.bas` flagged for rework; static spec documents flagged as out of date

## Session A — 26 June 2026

**Outcome:** `B3_Submissions.bas` written and tested. Orgs sheet population working against test database.

- `B3_Submissions.bas` written — new module, independent of existing modules
- `PopulateSubmissions` public Sub: calls API, iterates submission list, writes Org ID / Org Name / Submission Name / Submission ID to Orgs sheet in one block operation
- Reads `ProjectID`, `SubmissionYear`, `Toggle` from Config named ranges — no hardcoding
- Uses `Submissions` named range to locate header row; data written immediately below; 10,000-row clear before write
- `Option Private Module` identified as the cause of the macro not appearing in Excel's macro picker — removed
- Tested against test database with Managing Frailty (Project ID 35): 25 submissions returned correctly, first time
- No new architectural decisions — clean, self-contained deliverable

## Session B — 26 June 2026

**Outcome:** `Process_Folder.bas` written and tested. File picker, org/submission matching, and end-of-run summary all working.

- `Process_Folder.bas` written — new module, independent of existing modules
- Single button triggers multi-select `msoFileDialogFilePicker`; opens in folder stored in `SubmissionFolderPath`; writes folder of first selected file back to `SubmissionFolderPath` for next run
- `SubmissionFilePath` named range renamed to `SubmissionFolderPath` by user in `CCR_Tool_Base.xlsm`; Config sheet label updated to match
- `DataStart` and `DataMax` named ranges confirmed as already integrated in `B1_Importer.bas` — no changes needed
- Orgs sheet confirmed as populated with live Managing Frailty submission data from Session A test
- Matching decision tree implemented:
  - **Case 1** (no org match): message with filename and org name string; file skipped; no user choice
  - **Case 2** (one submission): Yes/No confirmation showing org name, submission name/ID, and template descriptor if present
  - **Case 3** (multiple submissions): numbered `InputBox` showing all submissions; confirmation message box under all input outcomes (valid, invalid, blank)
- `ProcessValidFile` stub confirms match with message box (filename + Submission ID)
- End-of-run summary: files selected / processed / skipped
- Tested against four files covering all three cases — all working first time, message boxes confirmed ideal
- **Architectural decision:** `Process_Folder` is a pre-step to the importer, not a replacement; `B1_Importer.bas` left untouched this session

## Design Session — 26 June 2026

**Outcome:** Module structure, separation of concerns, and session plan agreed. All documentation updated.

- Two problems identified and resolved:
  1. Module separation for validation concerns had not been fully thought through
  2. The data flow redesign (Sessions 10, A, B) had not been reflected in any documentation — static specs and dynamic docs all described the original single-file-path flow
- **Processing sequence agreed:** Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)
- **Module structure agreed:**
  - `B5_File_Validator` — structural and content validation of questionnaire files; runs before matching; no API calls
  - `B6_Response_Validator` — response text validation; orange cell colouring; local Drop downs lookup only
  - `B7_Duplicate_Detector` — duplicate detection via API; green cell colouring; separate module from B6
- **Managing Frailty established as primary build target**; Virtual Ward treated as a subsequent amendment session
- **Session plan agreed:** C (B5), D (wire B4→B5→B1, read-in test), E (create CCR records test), F (B6), G (B7), H (end-to-end test), I (Managing Frailty instance), J (Virtual Ward)
- All five documents updated: `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`, `Current_State.md`, `Next_Session.md`
- `Decisions.md` updated with new decisions: processing sequence, B1 parameter change, B6/B7 separation, Managing Frailty as primary target

## Session C — 26 June 2026

**Outcome:** `B5_File_Validator.bas` built and wired into `B4_Process_Folder.bas`. Validation pipeline ready to test. Key workbook constraint clarified.

- Validation design agreed collaboratively before any code was written:
  - **Check 1:** Mandatory sheets — `^`-delimited list in new Config named range `MandatorySheets`
  - **Check 2:** Support sheet field validation — XLookup for Project ID, Submission Period, Spec Type, Organisation Name; first two matched against Config named ranges, third against constant `"Clinical Case Review"`, fourth checked non-blank
  - **Check 3:** Spot checks — `^`-delimited list of `SheetName!CellRef:ExpectedValue` tokens in new Config named range `SpotChecks`; 10 checks targeting structural anchors across the Managing Frailty CCR sheet (first/last patient column headers, section headers, question type labels, fixed column headers)
- B5 receives an already-open `Workbook` object from B4 — file open/fail handled by B4 before B5 is called
- Each failed check produces a specific MsgBox and returns `False` to B4; B4 closes the file and skips
- `B4_Process_Folder.bas` updated:
  - Removed hardcoded `Support!B5` / `B6` cell reads and early Support sheet check
  - Calls `File_Validator.ValidateFile` immediately after open
  - Reads org name and submission descriptor via XLookup on Support sheet (consistent with B5 pattern)
  - `ProcessValidFile` stub comment updated to reflect current session
- `B3_Submissions.bas`: `Option Private Module` removed (had been missed from Session A tidy-up)
- Config sheet updated (by user directly in Excel):
  - Row 14: `Mandatory Sheets` / `CCR^Support` / named range `MandatorySheets`
  - Row 15: `Spot Checks` / 10-check string / named range `SpotChecks`
- **Constraint confirmed:** openpyxl drops Form Controls on save — `.xlsm` must never be passed through `write_excel` once buttons are present; all future workbook changes applied directly in Excel by user
- All `.bas` modules reimported by user; buttons reinstated
- Modules to reimport in Session D: `B1_Importer.bas`, `B4_Process_Folder.bas`

## Session D — 26 June 2026

**Outcome:** B1 updated and wired. Full pipeline confirmed working end-to-end. Module naming bug identified and fixed.

- `B1_Importer.bas` updated:
  - Now accepts `(Str_FilePath As String, Lng_SubmissionID As Long)` as parameters
  - `SubmissionFilePath` named range read removed
  - File path validation and sheet existence check removed (B5 owns these)
  - `FullDataArea.ClearContents` removed (moved to B4 at start of run)
  - Orientation `Else` guard removed
  - Paste row found dynamically — appends below last occupied row in column K (now J after column shift)
  - Submission ID written to column I on each imported row
  - Empty file check added after `Lng_StartIndex` resolved, before loop, for both orientations — probes first expected patient position; hard fail with explanatory MsgBox if blank
  - No end-of-run MsgBox — B4's summary covers that
- `B4_Process_Folder.bas` updated:
  - Clear prompt added at very top of `PickAndProcess` — Yes to clear `FullDataArea`, No to append
  - `ProcessValidFile` stub replaced with `Call B1_Importer.FileImporter(Str_FilePath, Lng_SubID)`
- **Module naming bug fixed:** both `B4_Process_Folder.bas` and `B5_File_Validator.bas` had incorrect `Attribute VB_Name` values (`"Process_Folder"` and `"File_Validator"` respectively). Both corrected to match their filenames. Cross-reference in B4 (`File_Validator.ValidateFile`) updated to `B5_File_Validator.ValidateFile`.
- All four modules (`B1`, `B4`, `B5` updated; `B5` name fix only) reimported by user
- Happy path confirmed working: file passes validation, matches to submission, rows imported correctly to Home sheet
- Representative failure cases confirmed: wrong project ID, missing sheet, failed spot check all produce correct MsgBox messages and skip correctly

## Session E — 26 June 2026

**Outcome:** Workbook fully configured for Managing Frailty. Home sheet columns extended. B1 and B4 updated. Test files populated. Session ordering corrected.

- **Base workbook reconfigured for Managing Frailty** (openpyxl rebuild from scratch — buttons and modules reinstated by user):
  - Home rows 2–6 fully populated: 38 questions, display sequence numbers, type codes (N/LS/YN/TX), source template row positions, real QIDs, short question labels
  - Named ranges `StartCols`, `QuestionCols`, `TypeCols`, `DataArea`, `FullDataArea` all extended from J:AV (was J:X)
  - Drop downs sheet populated from `managing_frailty_dropdowns.xlsx` with correct palette: row 1 METADATA yellow, row 2 HEADER_DARK navy, data rows no fill
  - Config credentials (B10/B11) cleared — previously contained plaintext credentials from earlier test session
- **Home sheet columns extended:** "Organisation name" column added at G; "Submission name" column inserted at H; existing Sub ID, CaseCode, Unique Ref shifted to I, J, K respectively. Excel managed all named range shifts automatically.
- **`B1_Importer.bas` updated:**
  - Signature extended: `FileImporter(Str_FilePath, Lng_SubmissionID, Str_OrgName, Str_SubName)`
  - Writes org name at `DataArea.Column - 4` (G), sub name at `DataArea.Column - 3` (H) on every imported row
  - Empty-record skip logic reworked: old check on unique ref (always populated in template) replaced with has-data check iterating all question positions in `StartCols` (index 2+); `Exit For` on first non-blank hit; `Lng_ImportedCount` tracks records; warns if zero records imported across whole file
- **`B4_Process_Folder.bas` updated:** both `FileImporter` call sites (Case 2 and Case 3) updated to pass `Str_OrgName` and the matched submission name
- **Test input files populated** with synthetic valid patient data via `populate_test_files.py` (written to `test_inputs/` and run directly by user): files 1 & 2 = 50 patients each, files 3 & 4 = 10 patients each; conditional logic respects blank rules (e.g. dementia timing only populated if dementia diagnosed)
- **Session ordering corrected:** B6_Response_Validator must precede the first live API test — invalid responses will cause the API post to fail. Session plan updated accordingly (was E=API test, F=B6; now F=B6, G=API test).
- Both modules reimported by user; confirmed working end-to-end with correct org name, sub name, and patient count on import

## Session F — 26 June 2026

**Outcome:** B6_Response_Validator built and passing. B1 and B4 updated. Error Log sheet added to workbook. Full import-and-validate pipeline confirmed working end-to-end.

- **A3 reviewed** before building B6 — confirmed ready for Managing Frailty with no changes needed. DT placeholder `"date"` left in place pending API team confirmation. The unconditional `Rng_Cell.Offset(0, -4).Value = "No"` on all rows noted as harmless and left untouched (Alex's pattern).
- **Functional spec corrected:** case code note field reference removed (column J receives the case code back from the API, not a unique reference string); response validation scope updated to current-run rows only; Error Log sheet added to outputs.
- **Error Log sheet** added to workbook by user (blank, row 1 bolded). Minimalist export design — no panel formatting.
- **`B1_Importer.bas` updated:** two `ByRef` output parameters added (`Lng_FirstRow`, `Lng_LastRow`) to enable B4 to track the run-wide row range. `Lng_FirstRow` set on first record written per call; `Lng_LastRow` updated on every record written.
- **`B4_Process_Folder.bas` updated:**
  - Run-wide row tracking variables added (`Lng_RunFirstRow`, `Lng_RunLastRow`, `Lng_FileFirstRow`, `Lng_FileLastRow`)
  - File-level variables reset to 0 before each B1 call; if B1 writes anything they update the run-level trackers
  - After file loop: if `Lng_RunFirstRow > 0`, calls `B6_Response_Validator.ValidateResponses`
  - Clear path changed from `.ClearContents` to `.Clear` (wipes formatting as well as values)
  - Grid borders reapplied after clear: four sides only (`xlEdgeTop`, `xlEdgeBottom`, `xlEdgeLeft`, `xlEdgeRight`, `xlInsideHorizontal`, `xlInsideVertical`) — diagonals explicitly excluded
- **`B6_Response_Validator.bas` built:** validates all five question types on current-run rows only:
  - LS: scans even column (QID col + 1) in Drop downs from row 3 downwards — bug found and fixed during testing (original code scanned odd column = list item IDs, not response text)
  - YN: exact match "Yes" or "No"
  - N: IsNumeric check
  - TX: IsNumeric check (must not be numeric)
  - DT: numeric serial within 46174–46269
  - Error Log: cleared and re-headed each run; one row per error (Row, Unique Ref, Question ID, Question No., Question Type, Invalid Value)
  - Summary MsgBox: count of errors found, or clean pass message
- All three modules reimported by user; confirmed passing with no errors on clean synthetic data

## Session G — 27 June 2026

**Outcome:** First live API test successful. Full end-to-end pipeline confirmed working. Multiple bugs identified and fixed.

- **Home sheet updated** by user before session start: new Org ID column inserted at column I; all named ranges updated in workbook accordingly.
- **`A3_API_Calls.bas` updated:**
  - `Rng_DataRows` changed from `Rng_FullDataArea.Columns(5)` to `Rng_FullDataArea.Columns(1).Cells` — iterates column F (Yes/No toggle) cell by cell
  - Loop condition changed from offset-based to `Rng_Cell.Cells(1).Value = "Yes"` — direct read of toggle cell
  - `Lng_OrgId` read from `Rng_Cell.Cells(1).Offset(0, 3)` (column I, Org ID) — `Split` on `"-"` removed
  - `Str_SubID` read from `Rng_Cell.Cells(1).Offset(0, 4)` (column J, Sub ID)
  - Both reads moved outside the question loop — once per row, not once per question
  - Case code writeback: `Rng_Cell.Cells(1).Offset(0, 5).Value = Str_CaseCode` (column K)
  - Yes→No: `If Rng_Cell.Cells(1).Value = "Yes" Then Rng_Cell.Cells(1).Value = "No"` — conditional, not unconditional
  - Exit For on blank cell added to avoid iterating entire FullDataArea
- **`B1_Importer.bas` updated:**
  - New parameter `ByVal Lng_OrgID As Long` added
  - `Rng_FullDataArea` declared and set
  - Writes `"Yes"` to `Rng_FullDataArea.Column` (column F) on each imported row
  - Writes `Lng_OrgID` to `DataArea.Column - 3` (column I) on each imported row
  - Column offsets updated throughout: OrgName → `-5`, SubName → `-4`, OrgID → `-3`, SubID → `-2`
- **`B4_Process_Folder.bas` updated:**
  - `Lng_OrgID2` / `Lng_OrgID3` declared and read from `Wsh_Orgs.Cells(row, 1)` at each match case
  - Both `FileImporter` call sites updated with new `Lng_OrgID` parameter
- **`A2_API_FUNCTIONS.bas` updated:**
  - Trailing comma bug fixed in `API_PostSurvey` payload builder: comma now prepended before each item (except first) using `Bln_FirstItem` flag, rather than appended after each item except last — eliminates trailing comma when last item has a blank value
  - Additionally: `Replace(Str_Payload & "]", ",]", "]")` added as a safety net
- **LS lookup bug fixed in A3:**
  - `Rng_LookupRange` now starts at row 3 (was row 2 — off by one; row 2 is the question label)
  - `Rng_LookupRange` now targets `Lng_QuestionLookupCol + 1` (even column = item text) and returns `Offset(0, -1)` (odd column = item ID) — previously had search and return columns swapped
- **For Each iteration instability:** user applied `.Cells` to `Rng_DataRows` declaration (`Rng_FullDataArea.Columns(1).Cells`) and `Rng_Cell.Cells(1)` throughout — resolves VBA range iteration instability on derived column ranges
- **API test results:** YN, N, TX all posting correctly first time. LS posting correctly after lookup fix. Case codes created, responses posted, case codes closed. Case code written back to column K. Yes flipped to No.
- All four modules (`A2`, `A3`, `B1`, `B4`) reimported by user after session fixes confirmed

## Session H — 27 June 2026

**Outcome:** B7_Duplicate_Detector built. Virtual Ward groundwork substantially complete. DT handling fully resolved. LS numeric matching fixed.

- **B7_Duplicate_Detector.bas** built:
  - Runs after B6, called from B4 with same `Lng_RunFirstRow` / `Lng_RunLastRow` parameters
  - Pure Home sheet logic — no API calls
  - For each new row, checks Sub ID (col J) + Unique Ref (col L) against all rows above `Lng_FirstRow`
  - On match: sets col F to No; colours all response cells green; runs cell-level comparison against matching earlier row; mismatches coloured orange (overrides green); logs duplicate record and each mismatch to Error Log
  - Error Log not cleared by B7 — B6 clears it; B7 appends
  - Summary MsgBox on completion

- **DT question handling fully resolved:**
  - `B6a_DT_Converter.bas` — new module; runs between B1 import and B6 validation
  - For each DT column: runs `TextToColumns` with `xlDMYFormat` on the imported row range to parse DD/MM/YYYY text as date serials; then iterates cells — formats each as Text (`@`) before writing `YYYY-MM-DD 00:00:00.000` string to prevent Excel re-interpreting on write-back
  - `B1_Importer.bas` reverted to clean original — DT values written as-is; B6a owns conversion
  - `B6_Response_Validator.bas` updated — DT validation now checks for `YYYY-MM-DD 00:00:00.000` format and date range rather than numeric serials; `IsValidDTString` private function added
  - `A3_API_Calls.bas` — DT case unchanged in behaviour; comment updated to clarify cell already holds formatted string from B6a

- **B4_Process_Folder.bas** updated:
  - Post-import call sequence: `B6a_DT_Converter` → `B6_Response_Validator` → `B7_Duplicate_Detector`
  - VW fallback lookup added for Support sheet — if "Submission Name" returns blank or "0", tries "Virtual Ward Name" (structural difference in VW template)

- **LS numeric matching fixed:**
  - B6: `Str_Response` now assigned with `CStr()` at point of read; Drop downs scan uses `Trim(CStr(...))` on lookup value
  - A3: LS case reads response with `CStr()` before XLookup
  - Root cause: Rockwood score question stores numeric list items (1–9); Excel converts these to numbers on import; string comparison was failing

- **Virtual Ward groundwork:**
  - VW_Data.xlsx produced with correct Config (Project ID 68, DataMax 75, SpotChecks for VW template), Home rows 2–6 (33 questions, correct types/QIDs/row positions/headers), Drop downs (18 LS questions, all list items)
  - Ethnicity column gap fixed (Mixed, Other, Unknown were missing)
  - Three VW test files generated via `populate_vw_test_files.py`: Essex Partnership (75 patients), Bromley Healthcare (10), Cornwall Partnership (10)
  - DropDownQs named range must be extended to `$A$1:$AI$1` in VW workbook instance

- **Open items carried forward:**
  - Drop downs numeric columns (Rockwood scores) must be formatted as Text in workbook before importing modules
  - VW workbook instance not yet built
  - API `questionType` string for DT still to be confirmed with API team
