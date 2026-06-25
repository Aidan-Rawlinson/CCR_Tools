# Architecture Design: CCR Tools

## Overview
A three-instance Excel/VBA tool sharing a single generalised VBA codebase. All project-specific variation is held in the workbook. The VBA reads configuration exclusively from named ranges; no project-specific values are hardcoded.

---

## Workbook Instances

| Instance | Project | Project ID | Status |
|---|---|---|---|
| `CCR_Tool_Base.xlsx` | — | — | Template; unpopulated |
| `CCR_Tool_ManagingFrailty.xlsx` | Managing Frailty in a Bed Based Setting | 35 | To be built |
| `CCR_Tool_VirtualWard.xlsx` | Virtual Wards (also known as Hospital at Home) | 68 | To be built |

Each instance is saved as `.xlsm` by the user after workbook creation, then `.bas` modules are imported via the VBA editor.

---

## Sheet Structure

| Sheet | Visible | Purpose |
|---|---|---|
| `Home` | Yes | Central working sheet — imported rows, toggle column, case code column, response columns |
| `Config` | Yes | All configuration — Project ID, Service ID, year, toggles, file path, credentials |
| `Orgs` | Yes | Organisation list and submission selector drop-downs |
| `Drop downs` | Yes | QID and list item ID lookup — keyed by question ID, built from SSMS CSVs |
| `Lists` | Hidden | Source lists for Config data validation drop-downs |

---

## Home Sheet Layout

Follows Alex's layout directly.

| Row | Content |
|---|---|
| 2 | Question numbers |
| 3 | Question type codes: `LS`, `YN`, `N`, `TX` |
| 4 | Question IDs — `QuestionCols` named range covers this row |
| 5 | Column headers |
| 6+ | Data rows (one row per patient) |

| Column | Named Range | Content |
|---|---|---|
| F | — | Import toggle — Yes / No |
| G | — | Source organisation |
| H | — | Source submission ID |
| I | — | Case code (written back after successful post) |
| J | — | Unique reference (Patient 1, Patient 2 etc.) |
| K+ | `QuestionCols` | Question response data |

---

## Config Sheet Layout

All configuration values in column B, labels in column A. Each cell is a named range. The VBA reads exclusively from these named ranges.

| Row | Named Range | Default | Input method |
|---|---|---|---|
| 2 | `ProjectID` | *(per instance)* | Plain value |
| 3 | `ServiceID` | `0` | Plain value |
| 4 | `SubmissionYear` | `2026` | Plain value |
| 5 | `DataSheetName` | `CCR` | Plain value |
| 6 | `Toggle` | `Test` | Data validation drop-down: `Test`, `Live` |
| 7 | `Orientation` | `Columns` | Data validation drop-down: `Columns`, `Rows` |
| 8 | `SubmissionFilePath` | *(blank)* | User entry at runtime |
| 10 | `APIUsername` | *(blank)* | User entry |
| 11 | `APIPassword` | *(blank)* | User entry |

---

## Drop Downs Sheet Layout

Follows Alex's structure. Question IDs in odd columns (row 1), question labels in row 2 (same odd columns), response text + list item ID pairs in the adjacent even column from row 3 downwards.

Built from SSMS CSVs per tool instance. Covers all `LS`-type questions. `TX`, `YN`, and `N` type questions do not require entries here.

---

## VBA Module Structure

| Module | Type | Purpose |
|---|---|---|
| `A1_API_SUPPORT` | Infrastructure | VBA-JSON library (Tim Hall, MIT), UTC utilities, `GetToken()` — reads credentials from `Config` named ranges |
| `A2_API_FUNCTIONS` | API layer | All API functions + `APICall` / `APIPost` HTTP wrappers — reads `Toggle` and `SubmissionYear` from `Config` |
| `A3_API_Calls` | Orchestration | `PostSurveyData` main loop — handles all four question types; reads `ServiceID` and `ProjectID` from `Config` |
| `B1_Importer` | Import | `FileImporter`, `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` — reads `Orientation`, `DataSheetName`, `SubmissionFilePath` from `Config` |
| `B2_Toggle` | UI | Environment toggle handler — writes to `Config!Toggle` |

---

## API Layer

### Base URLs
| Environment | Base URL |
|---|---|
| Live | `https://membersapi.nhsbenchmarking.nhs.uk/` |
| Test | `https://membersapidev.nhsbenchmarking.nhs.uk/` |

Selected by `Toggle` named range.

### API Calls

| Call | Method | Purpose |
|---|---|---|
| Get Submissions | GET | Retrieve submissions for selected org and project |
| Get Next Case Code | POST | Create next available case code; write unique reference to note field |
| Post Survey Data | POST | Submit response payload for a case code |
| Close Case Code | POST | Set case code status to Completed |
| Get Case Code Notes | GET | Retrieve existing case codes for duplicate detection |
| Get Case Code Responses | GET | Retrieve stored responses for colour-coded matching |

### Authentication
HTTP Basic Auth. Token fetched fresh per API call (no caching — follows Alex's pattern). Credentials read from `Config` named ranges `APIUsername` and `APIPassword`.

---

## Data Flow — Import

1. User pastes file path into `Config!SubmissionFilePath`
2. `FileImporter` opens the template file, reads sheet `DataSheetName`
3. `Orientation` toggle determines iteration direction:
   - `Columns`: iterate columns E+ (one patient per column); questions in rows
   - `Rows`: iterate rows 6+ (one patient per row); questions in columns (Alex's pattern)
4. Section header rows (no value in question type column) are skipped
5. Each patient's responses are written to a new row on `Home` starting at column K
6. `CaseCodeProcessed` calls Get Case Code Notes API; pre-marks rows with matching references as "No" in column F
7. `QuestionResponseMatcher` calls Get Case Code Responses API; colours cells green (match) or orange (mismatch/invalid)

## Data Flow — Post

For each row in Home where column F = "Yes":

1. Build response array — iterate `QuestionCols`, handle each question type:
   - `LS`: look up response text in `Drop downs` sheet → get list item ID
   - `YN`: convert "Yes" → `"Y"`, "No" → `"N"`
   - `N`: pass numeric value directly
   - `TX`: pass text value directly
2. Skip blank responses
3. Call Get Next Case Code → write unique reference to note field
4. Verify case code created
5. Post response array
6. Close case code
7. Write case code back to column I; set column F to "No"
