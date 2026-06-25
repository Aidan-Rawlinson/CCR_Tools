# Alex's Tool — Reference Document

> This document is a distilled reference for builders, not a full code listing.
> The source files in `reference/` are the ground truth. This document captures what is needed to make build decisions.

---

## Overview

A two-stage Excel/VBA tool. Stage 1 imports data from clinician-submitted Excel templates into a central `Home` sheet. Stage 2 posts that data to TBN's API record by record.

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| `Home` | Central data sheet — imported rows, import toggle column, case code column, question response columns |
| `Orgs` | Hidden/background sheet — holds the `Toggle` named range (Live/Test switch) |
| `Drop downs` | Lookup sheet — holds valid response options per question, keyed by question ID |

### Key Named Ranges

| Name | Sheet | Purpose |
|---|---|---|
| `Toggle` | `Orgs` | `"Live"` or `"Test"` — controls which API URLs are used |
| `QuestionCols` | `Home` | Range covering the question ID header row — used to iterate questions |
| `DropDownQs` | `Drop downs` | Header row of the dropdown lookup table — used to find the right column per question |
| `ProjectID` | `Home` | The project ID passed to API calls |
| `ServiceID` | `Home` | The service ID passed to `API_PostSurvey` |
| `SubmissionFolder` | `Home` | Folder path used by the file importer to find source templates |

### Home Sheet Layout (key columns)

| Column | Content |
|---|---|
| F | Import toggle — Yes/No dropdown (applied by importer) |
| G | Source organisation (populated by importer from source file `B5`) |
| H | Source submission ID (populated by importer from source file `B6`) |
| I | Case code (written back after successful API post) |
| J onwards | Question response data (pasted from source template row 11+) |
| Row 4 | Question IDs (used to build API payload) |
| Row 5 | Question type codes: `LS`, `YN`, or `N` (see Question Types below) |

---

## Source Template Structure

- Data sheet name: `Bed based CCR` *(project-specific — will differ in new tools)*
- Org identifier: cell `B5`
- Submission ID: cell `B6`
- Data rows start at: row 11
- Data columns: A–CJ (columns 1–88 scanned for last row detection)

---

## Question Types

| Code | Type | API type string | Notes |
|---|---|---|---|
| `LS` | List select | `"list"` | Response text looked up against `Drop downs` to get item ID |
| `YN` | Yes/No | `"yn"` | `"Yes"` → `"""Y"""`, `"No"` → `"""N"""` (JSON-quoted) |
| `N` | Numeric | `"number"` | Value passed directly |

---

## API Layer

### Authentication

- Endpoint: `GET https://membersapi.nhsbenchmarking.nhs.uk/authentication`
- Method: HTTP Basic Auth (username + password passed to `XMLHTTP.Open`)
- Returns: JSON `{ "data": { "token": "..." } }`
- Token is fetched fresh on every API call (no caching)
- Credentials in new tools: to be supplied via a dedicated input sheet *(not hardcoded)*

### Base URLs

All functions have a Live and Test variant, switched by the `Toggle` named range:

| Environment | Base |
|---|---|
| Live | `https://membersapi.nhsbenchmarking.nhs.uk/` |
| Test | `https://membersapidev.nhsbenchmarking.nhs.uk/` |

### API Calls

#### 1. Get Submissions
- **Method:** GET
- **URL:** `.../submissions/list?projectId=[ProjectID]&year=2026`
- **Returns:** Submission list filtered to the given org ID
- **Output:** Array of `[submissionId, submissionName]`
- **Note:** Year `2026` is hardcoded — will need to be parameterised in new tools

#### 2. Get Next Case Code
- **Method:** POST
- **URL:** `.../submissions/[SubmissionId]/addCnrCodes`
- **Payload:** `{"newCodeCount":1}`
- **Returns:** `data.newCnrCodes[1].caseCode`

#### 3. Post Survey Data
- **Method:** POST
- **URL:** `.../projects/questions?submissionId=[SubmissionId]&serviceId=[ServiceID]&submissionCaseCode=[CaseCode]`
- **Payload:** JSON array of question response objects:
  ```json
  [
    {"questionId":"[ID]","questionPart":1,"questionType":"[type]","value":[value]},
    ...
  ]
  ```
- **Success response:** `{"success":true}`
- **Note:** Rows with no response value are skipped (not included in payload)

#### 4. Close Case Code
- **Method:** POST
- **URL:** `.../submissions/[SubmissionId]/setCaseCodeCompleted`
- **Payload:** `{"caseCode": "[CaseCode]", "dataSubmitted": "Y"}`
- **Success response:** `{"success":true}`

#### 5. Get Case Code Notes (duplicate detection)
- **Method:** GET
- **URL:** `.../submissions/[SubmissionId]/caseCodes?allCaseCodes=true`
- **Returns:** Case codes where `dataSubmitted = "True"` and `completionStatus = "Completed"`, with embedded `externalCode` parsed from `caseCodeNotes` JSON string
- **Used by:** `CaseCodeProcessed` to pre-mark rows as already submitted (sets import toggle to "No")

#### 6. Get Case Code Responses (response matching)
- **Method:** GET
- **URL:** `.../projects/[ProjectId]/responses?year=2026&submissionId=[SubmissionId]`
- **Returns:** `[questionId, itemId]` pairs for a given case code
- **Used by:** `QuestionResponseMatcher` to colour-code cells green/orange vs existing data

---

## Data Flow — PostSurveyData

For each row in the data range where column F = "Yes":

1. If column I (case code) already has a value, prompt user to confirm re-import
2. Build responses array — iterate `QuestionCols`, handle each question type, skip blanks
3. Transpose array to avoid `Application.Transpose` single-row dimension bug *(Alex's workaround — manual loop into `Var_FinalArray`)*
4. Call **Get Next Case Code** → if blank, abort with error
5. Call **Post Survey Data** → if fails, abort with error
6. Call **Close Case Code** → if fails, abort with error
7. Write case code back to column I
8. Set column F to "No"

On completion: MsgBox "Import Complete"

---

## Module Summary

| Module | Type | Purpose |
|---|---|---|
| `A1_API_SUPPORT` | Infrastructure | VBA-JSON library (Tim Hall, MIT), UTC utilities, `GetToken()` |
| `A2_API_FUNCTIONS` | API layer | All six API functions plus `APICall` / `APIPost` HTTP wrappers |
| `A3_API_Calls` | Orchestration | `PostSurveyData` main loop + thin wrapper functions |
| `B1_Importer` | Import | `FileImporter`, `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` |
| `B2_Toggle` | UI | Live/Test toggle button handler |

---

## Known Gotchas

- **Hardcoded year:** `2026` appears in Get Submissions and Get Case Code Responses URLs — needs parameterising in new tools
- **Hardcoded sheet name:** `Bed based CCR` in `FileImporter` — project-specific, will differ per new tool
- **Hardcoded dev path:** `SaveJSONToFile` writes to `C:\Development Area\NACEL Submission Importer\` — dev artefact, not used in normal operation (call is commented out)
- **Transpose bug workaround:** When the responses array has only one row, `Application.Transpose` collapses it to 1D. Alex works around this with a manual loop into a fresh 2D array before passing to `API_PostSurvey`
- **Token fetched per call:** No token caching — each `APICall` / `APIPost` makes a fresh auth request first
- **Credentials:** Were hardcoded in `GetToken()` — replaced with `[USERNAME]` / `[PASSWORD]` placeholders in the `.bas` reference files and the `.xlsm`

---

## Guidance Document Notes

The PDF `Guidance(Different tool).pdf` in `reference/` is user-facing instructions for a **different project's clone** of this tool — same architecture, different API calls, different database setup. It is useful as a guide to the intended user journey but is not authoritative on implementation detail for the tools we are building.

### What it covers well
- The overall user journey and step sequence
- The colour-coding logic (orange = invalid response, green = matches database)
- The Yes/No column F gate for controlling which rows are imported

### Known discrepancies vs the code
- **File path vs folder path:** The instructions describe pasting a single file path; the code in `B1_Importer` reads a folder path and processes every `.xls*` file in it. The instructions likely describe an older or simpler version of the tool.
- **Live/Test toggle:** Not mentioned in the instructions at all. It exists in the code and is a significant operational control — users following the instructions alone would not know to check it. This mechanic will be carried forward into the new tools.
- **Project ID:** Instructions note it should be confirmed but don't explain that it is essentially fixed for a given tool instance.
- **Already-imported row detection:** Instructions treat this as a manual judgement call; the code does substantial work to assist (case code note matching, response comparison) but this isn't surfaced in the step-by-step guidance.
