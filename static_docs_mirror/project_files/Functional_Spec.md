# Functional Specification: CCR Tools

## Purpose
A bulk upload tool for Clinical Case Note Review (CCR) survey data. Clinicians complete an Excel template (one column per patient) and submit the file to TBN. The tool reads that file, validates the data, and posts each patient record to TBN's database via API.

## Scope
Two tools are in scope, one per project:
- **Managing Frailty in a Bed Based Setting** (Project ID 35)
- **Virtual Wards (also known as Hospital at Home)** (Project ID 68)

Both tools share identical functionality. They differ only in configuration (project ID, question set, org list, response lookups).

---

## User Journey

### 1. Select environment
The user confirms whether the tool is pointed at the Test or Live database via the Environment toggle on the Config sheet. Default is Test.

### 2. Enter Project ID and select organisation
The user confirms the Project ID (pre-populated, fixed per tool instance) and selects the submitting organisation from the drop-down on the Home sheet.

### 3. Retrieve submissions
The user clicks the Retrieve Submissions button. The tool calls the API to fetch all submissions for the selected organisation and populates a drop-down list. The user selects the correct submission.

### 4. Enter file path and import rows
The user pastes the file path of the received template into the Config sheet and clicks the Import button. The tool opens the template, reads all patient columns, and populates the Home sheet with one row per patient.

### 5. Review imported data
The tool automatically performs two checks and colour-codes the cells:
- **Orange:** the response text does not match any valid value in the Drop downs lookup sheet (invalid response)
- **Green:** a matching case code already exists in the database and the response matches what is stored (likely already imported)
- **No colour:** response is valid and not yet imported

The user reviews the imported rows. Orange cells must be corrected before import. Green rows should be deselected from import (column F set to "No") unless re-import is intended.

### 6. Select rows for import
The user confirms which rows to import by setting column F to "Yes" or "No". The tool pre-populates this column based on its duplicate detection logic, but the user has final control.

### 7. Import to database
The user clicks the Import Data to Database button. The tool processes each selected row sequentially:
1. Converts each response text to the corresponding list item ID (for list-select questions)
2. Builds the response payload
3. Creates a new case code via API, writing the unique reference ("Patient N") into the case code note field
4. Verifies the case code was created successfully
5. Posts the response data to the database via API
6. Sets the case code status to Completed

If any step fails, the tool reports the failure and stops processing that row. Subsequent rows are not affected.

On completion, a confirmation message is displayed.

---

## Inputs

| Input | Source | Notes |
|---|---|---|
| Submitted template file | File path entered by user | `.xlsx` format; sheet name `CCR`; questions in rows, patients in columns |
| Project ID | Config sheet (pre-populated) | Fixed per tool instance |
| Service ID | Config sheet (pre-populated) | 0 for both new projects |
| Submission year | Config sheet (pre-populated) | `2026` |
| Organisation | Drop-down on Home sheet | Drawn from Orgs sheet |
| Submission | Drop-down on Home sheet | Retrieved via API |
| Environment toggle | Config sheet | `Test` or `Live` |
| Credentials | Config sheet | Username and password for API authentication |

## Outputs

| Output | Destination | Notes |
|---|---|---|
| Case codes | TBN database (via API) | One per imported patient row |
| Survey responses | TBN database (via API) | One set per case code |
| Case code notes | TBN database (via API) | Unique reference ("Patient N") stored in note field |
| Import status | Home sheet column I | Case code written back after successful import |
| Import toggle | Home sheet column F | Set to "No" after successful import |

---

## Validation Rules

| Rule | Trigger | Response |
|---|---|---|
| Response text not in Drop downs | On import from template | Cell coloured orange; user must correct before posting |
| Case code already exists for this reference | On import from template | Row pre-set to "No"; cells coloured green if responses match |
| Case code creation failed | During database post | Error message shown; row skipped |
| Survey post failed | During database post | Error message shown; row skipped |

---

## Constraints
- Template files must be stored on a local drive (not a SharePoint URL)
- Maximum patients per file: 50 (Managing Frailty) / 75 (Virtual Ward)
- The tool cannot guarantee integrity of previously submitted rows if the template has been edited after submission
- Manual verification by the user is required before import
