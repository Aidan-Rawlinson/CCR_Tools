# Architecture Design: CCR Tools

## Overview
A three-instance Excel/VBA tool sharing a single generalised VBA codebase. All project-specific variation is held in the workbook. The VBA reads configuration exclusively from named ranges; no project-specific values are hardcoded.

---

## Workbook Instances

| Instance | Project | Project ID | Status |
|---|---|---|---|
| `CCR_Tool_Base.xlsm` | — | — | Template; base workbook built |
| `CCR_Tool_ManagingFrailty.xlsm` | Managing Frailty in a Bed Based Setting | 35 | To be built |
| `CCR_Tool_VirtualWard.xlsm` | Virtual Wards (also known as Hospital at Home) | 68 | To be built |

Each instance is created as `.xlsx` via the MCP write_excel tool, saved as `.xlsm` by the user in Excel, then `.bas` modules are imported via the VBA editor.

---

## Sheet Structure

| Sheet | Visible | Purpose |
|---|---|---|
| `Home` | Yes | Central working sheet — imported rows, toggle column, case code column, response columns |
| `Config` | Yes | All configuration — Project ID, Service ID, year, toggles, folder path, credentials |
| `Orgs` | Yes | Organisation and submission list — populated via API before file processing |
| `Drop downs` | Yes | QID and list item ID lookup — keyed by question ID, built from SSMS CSVs |
| `Lists` | Hidden | Source lists for Config data validation drop-downs |

---

## Home Sheet Layout

Follows Alex's layout directly.

| Row | Content |
|---|---|
| 2 | Question numbers |
| 3 | Question type codes: `LS`, `YN`, `N`, `TX`, `DT` |
| 4 | Question IDs and source positions — `QuestionCols` named range covers this row |
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
| 8 | `SubmissionFolderPath` | *(blank)* | Written by file picker at runtime |
| 10 | `APIUsername` | *(blank)* | User entry |
| 11 | `APIPassword` | *(blank)* | User entry |
| 12 | `DataStart` | *(per instance)* | First data column/row in source template |
| 13 | `DataMax` | *(per instance)* | Maximum number of patient records |

---

## Drop Downs Sheet Layout

Follows Alex's structure. Question IDs in odd columns (row 1), question labels in row 2 (same odd columns), response text + list item ID pairs in the adjacent even column from row 3 downwards.

Built from SSMS CSVs per tool instance. Covers all `LS`-type questions only. `TX`, `YN`, `N`, and `DT` type questions do not require entries here.

---

## VBA Module Structure

| Module | Purpose |
|---|---|
| `A1_API_SUPPORT` | VBA-JSON library (Tim Hall, MIT), UTC utilities, `GetToken()` — reads credentials from `Config` named ranges |
| `A2_API_FUNCTIONS` | All API functions plus `APICall` / `APIPost` HTTP wrappers — reads `Toggle` and `SubmissionYear` from `Config` |
| `A3_API_Calls` | `PostSurveyData` main loop — handles all five question types (LS, YN, N, TX, DT); reads `ServiceID` and `ProjectID` from `Config` |
| `B1_Importer` | Pure data transfer — reads the source template and writes patient rows to the Home sheet; accepts file path and submission ID as parameters; no validation logic |
| `B2_Toggle` | Environment toggle handler — writes to `Config!Toggle` |
| `B3_Submissions` | `PopulateSubmissions` — calls the API and writes org/submission data to the Orgs sheet |
| `B4_Process_Folder` | File picker and org/submission matching — presents the multi-select picker, calls B5 for each file, handles the matching decision tree, calls B1 once a match is confirmed |
| `B5_File_Validator` | File validation — determines whether a selected file is a structurally valid questionnaire file before matching proceeds; covers sheet presence, data layout, and configuration alignment |
| `B6_Response_Validator` | Response validation — runs after import; covers response text matching (orange cells), duplicate detection against the database (green cells), and case code comparison |

---

## Processing Flow

Files move through four stages in sequence. A file that fails at any stage does not proceed to the next.

```
Pick files (B4)
    → Validate file (B5)
        → Match to org/submission (B4)
            → Import (B1)
```

After all files are processed, response validation (B6) runs across all imported rows on the Home sheet.

Posting to the database (A3) is a separate user-initiated step after review.

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
| Get Submissions | GET | Retrieve all submissions for the project; used by B3 to populate Orgs sheet |
| Get Next Case Code | POST | Create next available case code; write unique reference to note field |
| Post Survey Data | POST | Submit response payload for a case code |
| Close Case Code | POST | Set case code status to Completed |
| Get Case Code Notes | GET | Retrieve existing case codes for duplicate detection |
| Get Case Code Responses | GET | Retrieve stored responses for colour-coded matching |

### Authentication
HTTP Basic Auth. Token fetched fresh per API call (no caching — follows Alex's pattern). Credentials read from `Config` named ranges `APIUsername` and `APIPassword`.

---

## Data Flow — File Processing

For each file selected via the picker:

1. **B5_File_Validator** checks the file is structurally valid
2. **B4_Process_Folder** reads org name and submission descriptor from `Support!B5` and `Support!B6`
3. Matching runs against the Orgs sheet; user confirms or selects submission
4. **B1_Importer** opens the file, reads sheet `DataSheetName`, iterates patient columns/rows per `Orientation`:
   - `Columns`: one patient per column; questions in rows; `DataStart` is a column letter
   - `Rows`: one patient per row; questions in columns; `DataStart` is a row number
5. Section header rows (no value in question type column) are skipped
6. Each patient's responses are written to a new row on `Home` starting at column K

After all files are processed:

7. **B6_Response_Validator** calls Get Case Code Notes API; pre-marks rows with matching references as "No" in column F
8. **B6_Response_Validator** calls Get Case Code Responses API; colours cells green (match) or orange (mismatch/invalid)

## Data Flow — Post

For each row in Home where column F = "Yes":

1. Build response array — iterate `QuestionCols`, handle each question type:
   - `LS`: look up response text in `Drop downs` sheet → get list item ID
   - `YN`: convert "Yes" → `"Y"`, "No" → `"N"`
   - `N`: pass numeric value directly
   - `TX`: pass text value directly
   - `DT`: validate date serial within range; convert to `YYYY-MM-DD 00:00:00.000`
2. Skip blank responses
3. Call Get Next Case Code → write unique reference to note field
4. Verify case code created
5. Post response array
6. Close case code
7. Write case code back to column I; set column F to "No"
