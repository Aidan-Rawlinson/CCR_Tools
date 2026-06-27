<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session F Complete — Ready for Session G (First Live API Test)

## Summary
Session F delivered B6_Response_Validator and the supporting changes to B1 and B4. The full import-and-validate pipeline is now working end-to-end. B6 validates all five question types (LS, YN, N, TX, DT), logs errors to a new Error Log sheet, and fires a summary message on completion. B4 was updated to track the run-wide row range and call B6 after all files are processed, and to use `.Clear` (not `.ClearContents`) with grid borders reapplied on the clear path. B1 was extended with two ByRef output parameters so B4 can track first/last row written. All changes tested and confirmed passing with no errors on clean data.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- `test_inputs/` — four non-broken Managing Frailty test files, each populated with synthetic valid patient data (files 1 & 2: 50 patients, files 3 & 4: 10 patients); four broken files left intentionally broken for B5 testing; `populate_test_files.py` script present
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsm` — fully configured for Managing Frailty; Error Log sheet added (blank, row 1 bolded); all `.bas` modules imported; buttons in place
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — built and tested
  - `A2_API_FUNCTIONS.bas` — built and tested
  - `A3_API_Calls.bas` — built; all five question types (LS, YN, N, TX, DT)
  - `B1_Importer.bas` — updated this session; two ByRef output parameters added (`Lng_FirstRow`, `Lng_LastRow`) so B4 can track the run-wide row range for B6
  - `B2_Toggle.bas` — built and tested
  - `B3_Submissions.bas` — built and tested against test database
  - `B4_Process_Folder.bas` — updated this session; tracks run-wide first/last row across all files; calls B6 after all files processed; clear path uses `.Clear` (wipes formatting as well as contents) then reapplies thin grid borders to FullDataArea (four sides only, diagonals excluded)
  - `B5_File_Validator.bas` — built and tested
  - `B6_Response_Validator.bas` — built and tested this session; validates all five question types; logs errors to Error Log sheet; summary MsgBox on completion

## Agreed module structure and processing sequence

Processing sequence per file: **Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)**

After all files processed:
- **Response validation (B6)** — runs on current-run rows only; orange cell colouring; logs to Error Log sheet
- **Duplicate detection (B7)** — not yet built; green cell colouring; API calls required

| Module | Status | Responsibility |
|---|---|---|
| `B1_Importer` | Updated and tested | Pure data transfer; ByRef row range output for B6 |
| `B4_Process_Folder` | Updated and tested | File picker, matching, calls B5/B1/B6; clear with border restore |
| `B5_File_Validator` | Built and tested | Structural and content validation of questionnaire files |
| `B6_Response_Validator` | Built and tested | Validates LS/YN/N/TX/DT; orange colouring; Error Log |
| `B7_Duplicate_Detector` | Not started | Duplicate detection via API; green cell colouring |

## B6 validation rules (confirmed)
| Type | Rule |
|---|---|
| `LS` | Response text must match a valid option in Drop downs (even column, rows 3+) |
| `YN` | Must be exactly "Yes" or "No" |
| `N` | Must be numeric |
| `TX` | Must not be numeric; any non-blank text string is valid |
| `DT` | Must be numeric Excel date serial within 46174–46269 (1 Jun–31 Aug 2026) |

## Home sheet column layout (current)
| Col | Content | Named range anchor |
|---|---|---|
| F | Process? | FullDataArea left edge |
| G | Organisation name | DataArea - 4 |
| H | Submission name | DataArea - 3 |
| I | Sub ID | DataArea - 2 |
| J | CaseCode | DataArea - 1 |
| K | Unique Ref. | DataArea (col 11) |
| L+ | Question responses | QuestionCols |

## Key constraint confirmed in Session C
**openpyxl drops Form Controls (buttons) on save.** The `.xlsm` must never be passed through the MCP `write_excel` tool once buttons are present. All future workbook changes are applied directly in Excel by the user, with Claude providing an exact instruction list.

## Agreed session plan

| Session | Goal |
|---|---|
| G | First live API test — Managing Frailty (gated on test database access) |
| H | Build B7_Duplicate_Detector |
| I | End-to-end test — Managing Frailty |
| J | Build Managing Frailty tool instance |
| K | Amend for Virtual Ward |

## Open items
- API `questionType` strings for `TX` and `DT` — placeholders in A3; needed before Virtual Ward end-to-end test (Session K); not blocking Sessions G–J
- Test database access — needed before Session G
