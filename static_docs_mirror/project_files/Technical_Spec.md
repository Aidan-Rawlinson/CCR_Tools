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
| `QuestionCols` | `Home!$K$4:$[last]$4` | QID header row — used to iterate questions during post |
| `StartCols` | `Home!$J$4:$[last]$4` | Source positions in template — unique ref plus all questions |
| `DropDownQs` | `'Drop downs'!$A$1:$[last]$1` | Header row of Drop downs lookup — used to find column per question |

---

## Key Technical Decisions

### Multi-select file picker replaces single file path
The original design used a single file path pasted by the user into Config. This has been replaced with a multi-select `msoFileDialogFilePicker`. The user selects one or more files via a dialog; no manual path entry is required. `SubmissionFolderPath` (formerly `SubmissionFilePath`) stores the last-used folder so the picker opens in the right place on subsequent runs.

### Processing sequence: validate → match → import
Each selected file passes through three stages before any data reaches the Home sheet. File validation (B5) runs first, in isolation, with no reference to the Orgs sheet. Org/submission matching (B4) runs only on files that pass validation. Import (B1) runs only on files that have been matched and confirmed. A file that fails at any stage is skipped with a message; subsequent files are not affected.

### B1_Importer accepts parameters, not named ranges
`FileImporter` accepts file path and submission ID as parameters passed by `B4_Process_Folder`. It does not read `SubmissionFolderPath` from the Config named range. All validation logic is removed — file validation is B5's responsibility.

### Response validation and duplicate detection in separate modules
Response text validation (B6) and duplicate detection (B7) are implemented in separate modules with distinct responsibilities. B6 is local only (Drop downs sheet lookup). B7 requires API calls. Both run after all files have been imported.

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
| `B1_Importer.bas` | Import | Needs update | To accept file path and submission ID as parameters; remove named range reads and validation logic |
| `B2_Toggle.bas` | UI | Built | Writes to Config!Toggle |
| `B3_Submissions.bas` | Submissions | Built | PopulateSubmissions tested against test database |
| `B4_Process_Folder.bas` | File picker / matching | Built | ProcessValidFile stub in place; B5 call and real B1 call not yet wired |
| `B5_File_Validator.bas` | File validation | Not started | |
| `B6_Response_Validator.bas` | Response validation | Not started | Orange cell colouring; local Drop downs lookup only |
| `B7_Duplicate_Detector.bas` | Duplicate detection | Not started | Green cell colouring; API calls required |

### Workbook Files (`.xlsx` → `.xlsm`)
| File | Purpose | Status |
|---|---|---|
| `CCR_Tool_Base.xlsm` | Base template — all sheets, Config populated with defaults, no project data | Built |
| `CCR_Tool_ManagingFrailty.xlsm` | Configured for Project 35 | Not started |
| `CCR_Tool_VirtualWard.xlsm` | Configured for Project 68 | Not started |

---

## Build Sequence

| Session | Deliverable |
|---|---|
| C | Build B5_File_Validator |
| D | Update B1_Importer to accept parameters; wire B4 → B5 → B1; test read-in of Managing Frailty templates |
| E | Test create CCR records from Managing Frailty files |
| F | Build B6_Response_Validator |
| G | Build and test B7_Duplicate_Detector |
| H | End-to-end test — Managing Frailty |
| I | Build Managing Frailty tool instance (Home sheet, named ranges, drop downs, import `.bas` files, workbook check) |
| J | Amend for Virtual Ward |

**Open items:** API `questionType` strings for `TX` and `DT` to be confirmed with API team before Session J.
