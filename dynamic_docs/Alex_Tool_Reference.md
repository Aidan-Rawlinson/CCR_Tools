# Alex's Tool — Reference Document

> This document is a distilled reference for builders, not a full code listing.
> The source files in `reference/` are the ground truth. This document captures what is needed to make build decisions.
>
> **Inspection history:** `.bas` modules read in Session 3 (original). `.xlsm` and `User_Template.xlsx` read in Session 3 (redo) using `read_excel`. Both files now fully inspected.
>
> **Known inspection gap:** `read_excel` (openpyxl) does not surface buttons — Form Controls and ActiveX Controls, including their labels, positions, and macro assignments, are invisible to this tool. Button details in this document were supplied by direct observation of the workbook by the user.

---

## Overview

A two-stage Excel/VBA tool. Stage 1 imports data from clinician-submitted Excel templates into a central `Home` sheet. Stage 2 posts that data to TBN's API record by record.

---

## Workbook Structure

| Sheet | Purpose |
|---|---|
| `Home` | Central data sheet — imported rows, import toggle column, case code column, question response columns |
| `Lists` | Small sheet (3 cells) — holds a "Submission Validation List" with 2 org entries. Likely a leftover from an earlier iteration; not active in current tool |
| `Drop downs` | Lookup sheet — holds valid response text and list item IDs per LS question, keyed by question ID |
| `Orgs` | Background sheet — holds the org list (29 orgs, a managed subset of the full programme) and the `Toggle` named range |

### Key Named Ranges

| Name | Points to | Value | Purpose |
|---|---|---|---|
| `Toggle` | `Orgs!J4` | `"Test"` | `"Live"` or `"Test"` — controls which API URLs are used |
| `QuestionCols` | `Home!$J$4:$CS$4` | — | QID header row — used to iterate questions during import |
| `DropDownQs` | `'Drop downs'!$A$1:$BX$1` | — | Header row of the dropdown lookup table — used to find the right column per question |
| `ProjectID` | `Home!$C$5` | `40` | Project ID passed to API calls |
| `ServiceID` | `Home!$C$7` | `146` | Service ID passed to `API_PostSurvey` |
| `SubmissionFolder` | `Home!$C$9` | *(dev path)* | Folder path used by the file importer. Contains a dev artefact path (`C:\Development Area\Member Submission Importer\Test Data\`) — must be replaced in any new tool instance |
| `Org_Id` | `Home!#REF!` | — | **Broken — points to a deleted or moved cell.** Code likely does not rely on this range; org ID comes from data rows directly |

### Home Sheet Layout

| Row | Content |
|---|---|
| 2 | Question numbers (1–86, with a labelling error — see Known Issues) |
| 3 | Question type codes: `YN`, `LS`, or `N` |
| 4 | Question IDs — header label "QID (hide)". This is the row `QuestionCols` iterates |
| 5 | Column headers: "Process?", "Organisation", "Sub ID", "CaseCode", "Unique Ref." in F–J; full question text from K onwards |
| 6+ | Data rows |

| Column | Content |
|---|---|
| C5 | `ProjectID` (value: 40) |
| C7 | `ServiceID` (value: 146) |
| C9 | `SubmissionFolder` (contains dev path — replace per instance) |
| F | Import toggle — Yes/No (applied by importer) |
| G | Source organisation (populated from template B5) |
| H | Source submission ID (populated from template B6) |
| I | Case code (written back after successful API post) |
| J | Unique reference number (from template column A) |
| K–CS | Question response data (88 questions; see Known Issues for anomalies) |

---

## Buttons

Buttons were not visible via `read_excel` inspection — details supplied by direct user observation. Three buttons exist on the Home sheet:

| Button label | Macro called | Purpose |
|---|---|---|
| *(unlabelled, next to Submission Folder Path)* | `FileImporter` | Reads the path in the Submission Folder Path cell, opens the template file(s) at that location, and pulls all patient records into the Home sheet |
| `DatabaseToggle` | `ToggleButton` | Switches the environment between Test and Live by writing to the `Toggle` named range |
| `Import Data to Database` | `PostSurveyData` | Iterates all rows in Home where column F = "Yes" and runs the full create/post/close API sequence for each |

**Note on `FileImporter` naming:** The button is positioned next to the file path input, which might suggest it refreshes the path. It does not — it triggers the full import of data from the template file at that path into the Home sheet. The name is accurate; the co-location with the path field is what causes the ambiguity.

---

## Source Template Structure (`User_Template.xlsx`)

- Data sheet name: `Bed based CCR` *(project-specific — will differ in new tools)*
- Service Item ID: cell `B1` (value: `146` — matches `ServiceID` in the tool; hidden row)
- Org identifier: cell `B5` *(blank — filled in by submitting trust; no validation in template)*
- Submission name: cell `B6` *(blank — filled in by submitting trust)*
- QID row: row 9 (hidden)
- Question text row: row 10
- Data rows start at: row 11
- Unique reference numbers: pre-populated in column A, rows 11–91, values 101–181
- **Fixed capacity: 81 rows.** Trusts submitting more than 81 records would need multiple files or a template extension
- Data columns: A–CJ (88 question columns)

### Template Sheet Structure

| Sheet | Purpose |
|---|---|
| `Bed based CCR` | Main data entry sheet — the one the importer reads |
| `Org list` | 95 organisations (full programme list, alphabetically sorted). Different from the tool's `Orgs` sheet which is a 29-org managed subset |
| `Drop downs` | Valid response text values only (no item IDs). This is what Excel's dropdown validation uses — distinct from the tool's `Drop downs` sheet which pairs text with list item IDs |

### Template Row 8 — Section Headers

Row 8 groups questions into named sections. Informational only — above the data range, not read by the importer:

| Section | Columns |
|---|---|
| Referral Information | B–G |
| Patient Information | H–M |
| Modified Barthel Index on admission | N–X |
| Record of staff contact | Y–AO |
| Care planning | AP–AW |
| Frailty | AX–BQ |
| Harm in care | BR–BS |
| Discharge information | BT–BY |
| Modified Barthel Index on discharge | BZ–CJ |

---

## Question Count

**Authoritative count: 88 questions** (confirmed from template QID row, columns A–CJ).

The tool's Home sheet question columns run K–CS (plus one anomalous column — see Known Issues). The `QuestionCols` named range covers `Home!$J$4:$CS$4`.

---

## Question Types

| Code | Type | API type string | Notes |
|---|---|---|---|
| `LS` | List select | `"list"` | Response text looked up against `Drop downs` to get item ID |
| `YN` | Yes/No | `"yn"` | `"Yes"` → `"""Y"""`, `"No"` → `"""N"""` (JSON-quoted) |
| `N` | Numeric | `"number"` | Value passed directly |

---

## Drop Downs Sheet

Structured with question IDs in row 1 (odd columns: A, C, E…), question labels in row 2 (same odd columns), and response text + list item ID pairs in the adjacent even column from row 3 downwards. Covers 38 LS-type questions. Not guaranteed to cover all LS questions — cross-reference against the full QID list before building.

The template's own `Drop downs` sheet contains the same response text values but without item IDs — that sheet drives Excel validation for clinicians. The tool's `Drop downs` sheet is the lookup used during import to resolve text → item ID.

---

## Orgs Sheet

29 organisations (a managed subset — trusts that agreed to submit by Excel). Columns: A = Org ID, B = Org Name, C = Concatenated display string. Submission ID column D is labelled but empty — submissions retrieved via API. `Toggle` named range at J4.

---

## API Layer

### Authentication

- Endpoint: `GET https://membersapi.nhsbenchmarking.nhs.uk/authentication`
- Method: HTTP Basic Auth (username + password passed to `XMLHTTP.Open`)
- Returns: JSON `{ "data": { "token": "..." } }`
- Token is fetched fresh on every API call (no caching)
- Credentials in new tools: to be supplied via a dedicated input sheet *(not hardcoded)*

### Base URLs

| Environment | Base |
|---|---|
| Live | `https://membersapi.nhsbenchmarking.nhs.uk/` |
| Test | `https://membersapidev.nhsbenchmarking.nhs.uk/` |

Switched by the `Toggle` named range.

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

## Known Issues and Anomalies

| Issue | Detail |
|---|---|
| **Question number labelling error** | Home row 2 skips a number — what would be Q51 ("If the patient received frailty screening, was the outcome added to their discharge summary?") has no label in row 2, but has a valid type code and QID. Data entry error in the workbook; not a structural issue |
| **BM column — database heading stored as question** | BM3 has type code `YN` but BM4 is empty and BM5 is the section header "55. Which of the following CGA components were completed during the admission?". The database treats headings as questions. No response data expected here; blank-skip logic should handle it gracefully — verify in Session 4 |
| **`Org_Id` named range broken** | Points to `Home!#REF!`. Likely a deleted or moved cell. Code does not appear to rely on it |
| **`SubmissionFolder` contains dev path** | `Home!C9` = `C:\Development Area\Member Submission Importer\Test Data\`. Must be replaced per tool instance |
| **Hardcoded year** | `2026` appears in Get Submissions and Get Case Code Responses URLs — needs parameterising in new tools |
| **Hardcoded sheet name** | `Bed based CCR` in `FileImporter` — project-specific, will differ per new tool |
| **Hardcoded dev path in SaveJSONToFile** | Writes to `C:\Development Area\NACEL Submission Importer\` — dev artefact, call is commented out |
| **Transpose bug workaround** | When responses array has only one row, `Application.Transpose` collapses to 1D. Alex works around this with a manual loop into a fresh 2D array |
| **Token fetched per call** | No token caching — each `APICall` / `APIPost` makes a fresh auth request first |
| **Credentials** | Were hardcoded in `GetToken()` — replaced with `[USERNAME]` / `[PASSWORD]` placeholders in reference files |
| **Buttons not visible via read_excel** | openpyxl does not surface Form Controls or ActiveX Controls. Button labels, positions, and macro assignments cannot be read programmatically — must be supplied by direct observation of the workbook |

---

## Guidance Document Notes

The file `Guidance(Different tool).docx` in `reference/` is user-facing instructions for a **different project's clone** of this tool — same architecture, different API calls, different database setup. Useful as a guide to the intended user journey but not authoritative on implementation detail for the tools we are building.

### What it covers well
- The overall user journey and step sequence
- The colour-coding logic (orange = invalid response, green = matches database)
- The Yes/No column F gate for controlling which rows are imported

### Known discrepancies vs the code
- **File path vs folder path:** Instructions describe pasting a single file path; the code reads a folder path and processes every `.xls*` file in it
- **Live/Test toggle:** Not mentioned in the instructions at all — a significant operational control users following the instructions alone would not know to check
- **Project ID:** Instructions note it should be confirmed but don't explain it is essentially fixed per tool instance
- **Already-imported row detection:** Instructions treat this as a manual judgement call; the code does substantial automated work to assist
