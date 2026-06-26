<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session C Complete — Ready for Session D (Wire B4→B5→B1, Read-in Test)

## Summary
Session C built `B5_File_Validator.bas` and wired it into `B4_Process_Folder.bas`. The full validation pipeline is now in place and ready to test at the start of Session D. A key workbook constraint was clarified: openpyxl drops Form Controls on save, so the `.xlsm` must never be modified via the MCP tool once buttons are present.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsm` — formatted to Alex's palette; all sheets, named ranges, and Config layout in place including new rows 14–15 (`MandatorySheets`, `SpotChecks`); buttons reinstated; all `.bas` modules imported
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — built and tested
  - `A2_API_FUNCTIONS.bas` — built and tested
  - `A3_API_Calls.bas` — built; all five question types (LS, YN, N, TX, DT)
  - `B1_Importer.bas` — built; needs update to accept parameters and remove validation logic (Session D)
  - `B2_Toggle.bas` — built and tested
  - `B3_Submissions.bas` — built and tested against test database
  - `B4_Process_Folder.bas` — updated this session; calls B5 before matching; reads org name and submission descriptor via XLookup; `ProcessValidFile` stub in place
  - `B5_File_Validator.bas` — built this session; three checks: mandatory sheets, Support sheet field validation, spot checks

## Agreed module structure and processing sequence

Processing sequence per file: **Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)**

After all files processed, two further steps run across all imported rows on the Home sheet:
- **Response validation (B6)** — orange cell colouring; local lookup against Drop downs sheet
- **Duplicate detection (B7)** — green cell colouring; API calls to retrieve existing case codes and responses

| Module | Status | Responsibility |
|---|---|---|
| `B1_Importer` | Needs update | Pure data transfer; accepts file path and submission ID as parameters; no validation logic |
| `B4_Process_Folder` | Updated — stub in place | File picker and org/submission matching; calls B5; reads Support via XLookup |
| `B5_File_Validator` | Built — untested | Structural and content validation of questionnaire files before matching |
| `B6_Response_Validator` | Not started | Response text validation; orange cell colouring; local Drop downs lookup only |
| `B7_Duplicate_Detector` | Not started | Duplicate detection via API; green cell colouring |

## Config sheet — current state
Rows 1–13: unchanged from previous session
Row 14: `MandatorySheets` — `CCR^Support`
Row 15: `SpotChecks` — 10 spot checks covering Managing Frailty CCR sheet structure

## Key constraint confirmed this session
**openpyxl drops Form Controls (buttons) on save.** The `.xlsm` must never be passed through the MCP `write_excel` tool once buttons are present. All future workbook changes are applied directly in Excel by the user, with Claude providing an exact instruction list.

## Agreed session plan

| Session | Goal |
|---|---|
| D | Update B1_Importer to accept parameters; wire B4 → B5 → B1; test read-in of Managing Frailty templates |
| E | Test create CCR records from Managing Frailty files |
| F | Build B6_Response_Validator |
| G | Build and test B7_Duplicate_Detector |
| H | End-to-end test — Managing Frailty |
| I | Build Managing Frailty tool instance (Home sheet, named ranges, drop downs, import `.bas` files, workbook check) |
| J | Amend for Virtual Ward |

## Open items
- API `questionType` strings for `TX` and `DT` — placeholders in A3; needed before Virtual Ward end-to-end test (Session J); not blocking Sessions D–I
- `B1_Importer.bas` still reads `SubmissionFilePath` from named range and contains validation logic — both to be removed in Session D
- **Flag for Session D:** empty file check (no patient data in submitted template) to be built into B1_Importer as a hard fail at read-in time
