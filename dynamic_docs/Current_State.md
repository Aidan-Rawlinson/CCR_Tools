<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session G Complete — Upload Confirmed Fully Working

## Summary
Session G delivered the first successful end-to-end API test. All question types (YN, N, TX, LS) posting correctly to the test database. Case codes created, responses posted, case codes closed, case code written back to the Home sheet, Yes flipped to No. Several bugs identified and fixed during the session: column layout changes (new Org ID column), B1 not writing Yes or Org ID on import, A3 iterating wrong column and using wrong offsets, trailing comma in JSON payload, LS lookup scanning wrong column and direction, For Each iteration instability resolved with .Cells. All fixes applied and confirmed working.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- `test_inputs/` — four non-broken Managing Frailty test files, each populated with synthetic valid patient data (files 1 & 2: 50 patients, files 3 & 4: 10 patients); four broken files left intentionally broken for B5 testing; `populate_test_files.py` script present
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsm` — fully configured for Managing Frailty; Error Log sheet added; all `.bas` modules imported; buttons in place; Org ID column added to Home sheet
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — built and tested
  - `A2_API_FUNCTIONS.bas` — updated this session; trailing comma bug fixed in `API_PostSurvey` payload builder
  - `A3_API_Calls.bas` — updated this session; iterates `FullDataArea.Columns(1).Cells` (column F); correct offsets for new column layout; Org ID and Sub ID read outside question loop; LS lookup fixed
  - `B1_Importer.bas` — updated this session; writes "Yes" to column F and Org ID to column I on each imported row; new `Lng_OrgID` parameter; updated column offsets for new layout
  - `B2_Toggle.bas` — built and tested
  - `B3_Submissions.bas` — built and tested
  - `B4_Process_Folder.bas` — updated this session; reads Org ID from Orgs sheet column 1; passes to B1
  - `B5_File_Validator.bas` — built and tested
  - `B6_Response_Validator.bas` — built and tested

## Home sheet column layout (current)
| Col | Content | Notes |
|---|---|---|
| F | Process? (Yes/No) | FullDataArea col 1; written by B1 on import |
| G | Organisation name | DataArea - 5 |
| H | Submission name | DataArea - 4 |
| I | Org ID | DataArea - 3; new this session |
| J | Sub ID | DataArea - 2 |
| K | CaseCode | DataArea - 1; written back by A3 after successful post |
| L | Unique Ref. | DataArea (col 1) |
| M+ | Question responses | QuestionCols |

## Offsets from column F (FullDataArea col 1) used in A3
| Offset | Column | Content |
|---|---|---|
| +3 | I | Org ID |
| +4 | J | Sub ID |
| +5 | K | CaseCode |

## Agreed module structure and processing sequence

Processing sequence per file: **Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)**

After all files processed:
- **Response validation (B6)** — runs on current-run rows only; orange cell colouring; logs to Error Log sheet
- **Post to database (A3)** — confirmed working end-to-end this session
- **Duplicate detection (B7)** — not yet built; green cell colouring; API calls required

## Key constraint confirmed in Session C
**openpyxl drops Form Controls (buttons) on save.** The `.xlsm` must never be passed through the MCP `write_excel` tool once buttons are present.

## Agreed session plan

| Session | Goal |
|---|---|
| H | Build B7_Duplicate_Detector |
| I | End-to-end test — Managing Frailty |
| J | Build Managing Frailty tool instance |
| K | Amend for Virtual Ward |

## Open items
- API `questionType` strings for `TX` and `DT` — `"text"` and `"date"` used as placeholders; TX confirmed working this session; DT still to be confirmed before Virtual Ward
- B7_Duplicate_Detector not yet built — green cell colouring; API calls required
