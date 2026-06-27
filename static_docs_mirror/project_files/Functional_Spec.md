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

### 2. Populate organisation and submission list
The user clicks the Populate Submissions button. The tool calls the API and populates the Orgs sheet with all organisations and submissions for the project. This must be done before files can be processed.

### 3. Select and process files
The user clicks the Process Files button. A multi-select file picker opens, defaulting to the folder used on the previous run. The user selects one or more submitted template files.

For each selected file, the tool runs the following sequence automatically:

**a. File validation**
The tool checks that the file is a valid questionnaire file — correct sheet present, data in the expected structure, configuration alignment confirmed. Files that fail validation are skipped with a message; the user does not proceed to matching for an invalid file.

**b. Org/submission matching**
The tool reads the organisation name and submission descriptor from the file and attempts to match against the populated Orgs sheet:
- **No match:** file is skipped with a message
- **One match:** user is shown the match and asked to confirm before proceeding
- **Multiple matches:** user is shown all matches and asked to select the correct one before proceeding

**c. Import**
Once matched and confirmed, the tool imports the file: reads all patient columns from the template and writes one row per patient to the Home sheet.

### 4. Review imported data
After all files have been processed, the tool performs response validation on the rows just imported and colour-codes any invalid cells:
- **Orange:** the response text does not match any valid value in the Drop downs lookup sheet (invalid response); must be corrected before posting
- **No colour:** response is valid and ready to post

Response validation runs only on rows imported in the current run. Rows already present on the Home sheet from a previous run are not re-validated.

An Error Log sheet records the detail of each validation failure — row number, unique reference, question ID, and the invalid value. This log is cleared at the start of each validation run. A summary message is shown on completion indicating how many errors were found.

The user reviews any orange cells and corrects them before proceeding. Green cell colouring (duplicate detection) is handled separately by B7 and is not part of this step.

### 5. Select rows for import
The user confirms which rows to post by setting column F to "Yes" or "No". The tool pre-populates this column based on duplicate detection (B7), but the user has final control.

### 6. Post to database
The user clicks the Import Data to Database button. The tool processes each selected row sequentially:
1. Converts each response text to the corresponding list item ID (for list-select questions)
2. Builds the response payload
3. Creates a new case code via API
4. Verifies the case code was created successfully
5. Posts the response data to the database via API
6. Sets the case code status to Completed
7. Writes the case code back to column J on the Home sheet; sets column F to "No"

If any step fails, the tool reports the failure and stops processing that row. Subsequent rows are not affected.

On completion, a confirmation message is displayed.

---

## Inputs

| Input | Source | Notes |
|---|---|---|
| Submitted template files | Multi-select file picker | `.xlsx` format; sheet name `CCR`; questions in rows, patients in columns |
| Project ID | Config sheet (pre-populated) | Fixed per tool instance |
| Service ID | Config sheet (pre-populated) | 0 for both new projects |
| Submission year | Config sheet (pre-populated) | `2026` |
| Environment toggle | Config sheet | `Test` or `Live` |
| Credentials | Config sheet | Username and password for API authentication |

## Outputs

| Output | Destination | Notes |
|---|---|---|
| Case codes | TBN database (via API) | One per imported patient row |
| Survey responses | TBN database (via API) | One set per case code |
| Import status | Home sheet column J | Case code written back after successful post |
| Import toggle | Home sheet column F | Set to "No" after successful post |
| Validation errors | Error Log sheet | Row, unique reference, question ID, invalid value; cleared at start of each validation run |

---

## Validation Rules

### File validation (before import)
| Rule | Response |
|---|---|
| File cannot be opened | File skipped with message |
| Expected sheet not present | File skipped with message |
| Data not in expected structure or position | File skipped with message |

### Response validation (after import — current run rows only)
| Rule | Trigger | Response |
|---|---|---|
| LS response text not in Drop downs | After import completes | Cell coloured orange; error logged; user must correct before posting |
| N response is not numeric | After import completes | Cell coloured orange; error logged; user must correct before posting |

Note: YN and TX responses are not validated. YN values are constrained by template drop-downs. TX accepts any non-blank text; blank TX cells are skipped at post time.

### Post validation (during database post)
| Rule | Response |
|---|---|
| Case code creation failed | Error message shown; row skipped |
| Survey post failed | Error message shown; row skipped |

---

## Constraints
- Template files must be stored on a local drive (not a SharePoint URL)
- Maximum patients per file: 50 (Managing Frailty) / 75 (Virtual Ward)
- The tool cannot guarantee integrity of previously submitted rows if the template has been edited after submission
- Manual verification by the user is required before posting
