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
| `TX` | TBC | To be confirmed with API team before Session 9 |

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
| `SubmissionFilePath` | `Config!$B$8` | Path to submitted template file |
| `APIUsername` | `Config!$B$10` | API authentication username |
| `APIPassword` | `Config!$B$11` | API authentication password |
| `QuestionCols` | `Home!$K$4:$[last]$4` | QID header row — used to iterate questions during post |
| `DropDownQs` | `'Drop downs'!$A$1:$[last]$1` | Header row of Drop downs lookup — used to find column per question |

---

## Key Technical Decisions

### Transpose bug workaround
Alex's tool works around a VBA `Application.Transpose` limitation (collapses single-row arrays to 1D) using a manual loop into a fresh 2D array. This workaround is carried forward unchanged.

### Token per call
No token caching. Each `APICall` / `APIPost` makes a fresh auth request. Follows Alex's pattern — acceptable for the volume of calls involved.

### Sequential row processing
Rows are processed one at a time, completing the full create/post/close sequence before moving to the next. If a row fails, processing stops for that row and continues with the next. Follows Alex's pattern.

### Single file path, not folder
Alex's tool reads a folder path and processes every `.xls*` file in it. The new tools use a single file path (held in `Config!SubmissionFilePath`). This is simpler, less error-prone, and matches the actual workflow (one file per submission batch).

### Section header skip
The new templates include section header rows in the data range (rows with no question type value in column D). The importer skips any row where the question type cell is blank.

### Blank response skip
Rows with no response value are not included in the API payload. Follows Alex's pattern. Applies to all question types including `TX`.

---

## Source Files

### VBA Modules (`.bas`)
| File | Module | Notes |
|---|---|---|
| `A1_API_SUPPORT.bas` | Infrastructure | Credentials now read from Config, not hardcoded |
| `A2_API_FUNCTIONS.bas` | API layer | `SubmissionYear` read from Config, not hardcoded |
| `A3_API_Calls.bas` | Orchestration | Handles four question types; reads ServiceID from Config |
| `B1_Importer.bas` | Import | Orientation-aware; reads DataSheetName from Config |
| `B2_Toggle.bas` | UI | Writes to Config!Toggle |

### Workbook Files (`.xlsx` → `.xlsm`)
| File | Purpose |
|---|---|
| `CCR_Tool_Base.xlsx` | Base template — all sheets, Config populated with defaults, no project data |
| `CCR_Tool_ManagingFrailty.xlsx` | Configured for Project 35 |
| `CCR_Tool_VirtualWard.xlsx` | Configured for Project 68 |

---

## Build Sequence

| Session | Deliverable |
|---|---|
| 6 | SSMS CSVs supplied; Drop downs sheets built; base workbook structure created |
| 7 | `A1_API_SUPPORT.bas`, `A2_API_FUNCTIONS.bas` written; API layer tested against test DB |
| 8 | `B1_Importer.bas` written; orientation toggle; four question types; section header skip |
| 9 | `A3_API_Calls.bas`, `B2_Toggle.bas` written; end-to-end test on test DB |
| 10 | Managing Frailty instance configured; `.bas` files imported; full test |
| 11 | Virtual Ward instance configured; `.bas` files imported; full test; Alex handoff notes |

**Gate:** Test database access must be confirmed before Session 7 begins.
