# Technical Specification: CCR Tools

## Platform
- **Language:** VBA (Excel macro-enabled workbook)
- **Host application:** Microsoft Excel (Windows)
- **File format:** `.xlsx` created via openpyxl; saved as `.xlsm` by user before VBA import
- **VBA modules:** Exported as `.bas` files; imported manually via VBA editor

## Environment
- **Development machine:** Windows (TBN staff machine)
- **MCP server:** Python, running on same Windows machine
- **openpyxl version:** 3.1.5 (confirmed installed)
- **MCP tools available:** `read_excel`, `write_excel`, `read_file`, `write_file`, `git_commit`, `git_revert`

---

## Dependencies

### VBA
- **VBA-JSON** (Tim Hall, MIT licence) — bundled in `A1_API_SUPPORT.bas`; handles JSON parsing and serialisation
- **MSXML2.XMLHTTP** — Windows built-in; used for all HTTP calls
- **Microsoft Scripting Runtime** — for `Dictionary` object used in response array construction

### Python (MCP server)
- **openpyxl 3.1.5** — workbook creation and population
- **mcp** — MCP server framework

---

## API

### Authentication
- Method: HTTP Basic Auth
- Endpoint: `GET https://membersapi.nhsbenchmarking.nhs.uk/authentication`
- Returns: JSON `{ "data": { "token": "..." } }`
- Token fetched fresh on every API call (no caching — follows Alex's pattern)
- Credentials source: `Config` sheet named ranges `APIUsername` / `APIPassword`

### Environments
| Environment | Base URL |
|---|---|
| Live | `https://membersapi.nhsbenchmarking.nhs.uk/` |
| Test | `https://membersapidev.nhsbenchmarking.nhs.uk/` |

### Question Type Strings (API payload)
| Config code | API `questionType` string | Notes |
|---|---|---|
| `LS` | `"list"` | Confirmed from Alex's tool |
| `YN` | `"yn"` | Confirmed from Alex's tool |
| `N` | `"number"` | Confirmed from Alex's tool |
| `TX` | TBC | To be confirmed with API team |
| `DT` | TBC | To be confirmed with API team — Virtual Ward only |

---

## Named Ranges

All named ranges are workbook-scoped and point to the `Config` sheet unless otherwise noted.

| Name | Points to | Purpose |
|---|---|---|
| `ProjectID` | `Config!$B$2` | Project ID passed to API calls |
| `ServiceID` | `Config!$B$3` | Service ID passed to Post Survey Data |
| `SubmissionYear` | `Config!$B$4` | Year parameter in API URLs |
| `DataSheetName` | `Config!$B$5` | Sheet name to read in submitted template |
| `Toggle` | `Config!$B$6` | `Test` or `Live` — controls API base URL |
| `Orientation` | `Config!$B$7` | `Columns` or `Rows` — controls importer iteration |
| `SubmissionFolderPath` | `Config!$B$8` | Last-used folder path; written by file picker at runtime |
| `APIUsername` | `Config!$B$10` | API authentication username |
| `APIPassword` | `Config!$B$11` | API authentication password |
| `DataStart` | `Config!$B$12` | First data column letter (Columns) or row number (Rows) in source template |
| `DataMax` | `Config!$B$13` | Maximum number of patient records per file |
| `MandatorySheets` | `Config!$B$14` | `^`-delimited list of sheet names required in submitted files |
| `SpotChecks` | `Config!$B$15` | `^`-delimited list of `SheetName!CellRef:ExpectedValue` structural checks |
| `TypeCols` | `Home!$K$3:$AV$3` | Question type codes — used by B6 to determine validation rule per column |
| `StartCols` | `Home!$K$4:$AV$4` | Source positions in template — unique ref plus all questions |
| `QuestionCols` | `Home!$K$5:$AV$5` | QID row — used to iterate questions during post and B6 LS lookup |
| `DataArea` | `Home!$K$7:$AV$19408` | Data area anchor; column K is unique ref; questions from L onwards |
| `FullDataArea` | `Home!$F$7:$AV$19408` | Full data area including process toggle and org/sub columns |
| `DropDownQs` | `'Drop downs'!$A$1:$S$1` | QID header row of Drop downs sheet — used by B6 to locate column per question |
| `Submissions` | `Orgs!$A$1:$D$1` | Header row anchor for Orgs sheet |

---

## Home Sheet Column Layout

| Column | Content | VBA reference |
|---|---|---|
| F | Process? (Yes/No) | `FullDataArea` left edge |
| G | Organisation name | `DataArea.Column - 4` |
| H | Submission name | `DataArea.Column - 3` |
| I | Sub ID | `DataArea.Column - 2` |
| J | CaseCode | `DataArea.Column - 1` |
| K | Unique Ref. | `DataArea.Column` (col 11) |
| L+ | Question responses | `QuestionCols` |

---

## Key Technical Decisions

### Multi-select file picker replaces single file path
The original design used a single file path pasted by the user into Config. This has been replaced with a multi-select `msoFileDialogFilePicker`. The user selects one or more files via a dialog; no manual path entry is required. `SubmissionFolderPath` (formerly `SubmissionFilePath`) stores the last-used folder so the picker opens in the right place on subsequent runs.

### Processing sequence: validate → match → import
Each selected file passes through three stages before any data reaches the Home sheet. File validation (B5) runs first, in isolation, with no reference to the Orgs sheet. Org/submission matching (B4) runs only on files that pass validation. Import (B1) runs only on files that have been matched and confirmed. A file that fails at any stage is skipped with a message; subsequent files are not affected.

### B1_Importer accepts parameters, not named ranges
`FileImporter` accepts file path, submission ID, org name, and submission name as parameters passed by `B4_Process_Folder`. It does not read these values from named ranges. All validation logic is removed — file validation is B5's responsibility.

### B1_Importer writes org name and submission name on import
On each imported row, B1 writes org name to `DataArea.Column - 4` (G) and submission name to `DataArea.Column - 3` (H). Both values are passed in as parameters from B4, which reads them from the matched submission in the Orgs sheet.

### B1 empty-record skip: has-data check across question cells
Patient positions are skipped if all question cells are blank. The threshold is 1: at least one non-blank response is required to import a record. The unique reference row ("Patient 1" etc.) is not used for this check as it is hardcoded in the template and is never blank. The check iterates `StartCols` positions from index 2 onwards (index 1 is the unique ref) and exits on first non-blank hit.

### Response validation and duplicate detection in separate modules
Response text validation (B6) and duplicate detection (B7) are implemented in separate modules with distinct responsibilities. B6 is local only (Drop downs sheet lookup, numeric check). B7 requires API calls. Both run after all files have been imported. B6 must run before the API post.

### B6 must precede the first API post
The API import will fail if responses are invalid. B6 flags invalid cells orange before any posting occurs, giving the user the opportunity to correct data. Running a post against unvalidated data risks creating partial or corrupt case codes.

### Submissions populated via API, not user-selected drop-down
The Orgs sheet is populated by `B3_Submissions.PopulateSubmissions` via an API call before file processing begins.

### Transpose bug workaround
Alex's tool works around a VBA `Application.Transpose` limitation (collapses single-row arrays to 1D) using a manual loop into a fresh 2D array. This workaround is carried forward unchanged.

### Token per call
No token caching. Each `APICall` / `APIPost` makes a fresh auth request. Follows Alex's pattern.

### Sequential row processing
Rows are processed one at a time, completing the full create/post/close sequence before moving to the next. Follows Alex's pattern.

### Section header skip
The new templates include section header rows in the data range (rows with no question type value in column D). The importer skips any row where the question type cell is blank.

### Blank response skip
Rows with no response value are not included in the API payload. Applies to all question types including `TX` and `DT`.

### DT question handling
Virtual Ward only. Users enter dates in `DD/MM/YYYY` format but Excel silently converts these to date serial numbers. If numeric, validate within 1 June–31 August 2026 (serials 46,174–46,269); if valid, convert to `YYYY-MM-DD 00:00:00.000`. Non-numeric or out-of-range values flagged orange. API `questionType` string to be confirmed.

---

## Source Files

### VBA Modules (`.bas`)
| File | Module | Status | Notes |
|---|---|---|---|
| `A1_API_SUPPORT.bas` | Infrastructure | Built | Credentials read from Config named ranges |
| `A2_API_FUNCTIONS.bas` | API layer | Built | `SubmissionYear` read from Config |
| `A3_API_Calls.bas` | Orchestration | Built | Handles five question types; reads ServiceID from Config |
| `B1_Importer.bas` | Import | Built | Accepts file path, submission ID, org name, sub name as parameters; writes all four to Home; has-data skip logic |
| `B2_Toggle.bas` | UI | Built | Writes to Config!Toggle |
| `B3_Submissions.bas` | Submissions | Built | PopulateSubmissions tested against test database |
| `B4_Process_Folder.bas` | File picker / matching | Built | Calls B5, reads org/sub from Orgs sheet, calls B1 with all four parameters |
| `B5_File_Validator.bas` | File validation | Built | Config-driven mandatory sheet and spot check validation |
| `B6_Response_Validator.bas` | Response validation | Not started | LS lookup, N numeric check; orange cell colouring; must run before API post |
| `B7_Duplicate_Detector.bas` | Duplicate detection | Not started | Green cell colouring; API calls required |

### Workbook Files (`.xlsx` → `.xlsm`)
| File | Purpose | Status |
|---|---|---|
| `CCR_Tool_Base.xlsm` | Fully configured for Managing Frailty — all sheets, named ranges, Home rows 2–6, Drop downs populated | Built |
| `CCR_Tool_ManagingFrailty.xlsm` | Separate instance for Project 35 | Not started |
| `CCR_Tool_VirtualWard.xlsm` | Configured for Project 68 | Not started |

---

## Build Sequence

| Session | Deliverable |
|---|---|
| C | Build B5_File_Validator |
| D | Update B1_Importer; wire B4 → B5 → B1; test read-in of Managing Frailty templates |
| E | Workbook configured for Managing Frailty; Home columns extended; B1/B4 updated; test files populated |
| F | Build B6_Response_Validator |
| G | First live API test — Managing Frailty (gated on test database access) |
| H | Build B7_Duplicate_Detector |
| I | End-to-end test — Managing Frailty |
| J | Build Managing Frailty tool instance |
| K | Amend for Virtual Ward |

**Open items:** API `questionType` strings for `TX` and `DT` to be confirmed with API team before Session K.
