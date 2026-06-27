<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session H Complete — Virtual Ward build substantially complete; DT handling resolved

## Summary
Session H delivered B7_Duplicate_Detector, substantial Virtual Ward workbook groundwork, and a full DT question handling solution. The processing pipeline now includes B6a_DT_Converter between import and validation. B4 updated with Virtual Ward submission name fallback. LS numeric matching fixed in B6 and A3 via CStr(). VW_Data.xlsx produced with correct Config, Home rows 2–6, and Drop downs content. Three VW test files generated. Outstanding: Drop downs sheet numeric columns need Text formatting; VW workbook instance not yet built.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- `test_inputs/` — four non-broken Managing Frailty test files; four broken files for B5 testing; `populate_test_files.py`; three VW test files (Essex 75 patients, Bromley 10, Cornwall 10); `populate_vw_test_files.py`
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsm` — fully configured for Managing Frailty; all modules imported including B6a and B7; buttons in place
  - `VW_Data.xlsx` — Config, Home rows 2–6, Drop downs for Virtual Ward; to be lifted across to VW workbook instance
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — built and tested
  - `A2_API_FUNCTIONS.bas` — built and tested
  - `A3_API_Calls.bas` — updated this session; LS case uses CStr() on response value before XLookup; DT case passes formatted string directly
  - `B1_Importer.bas` — reverted to clean original; DT handling removed (now owned by B6a)
  - `B2_Toggle.bas` — built and tested
  - `B3_Submissions.bas` — built and tested
  - `B4_Process_Folder.bas` — updated this session; calls B6a → B6 → B7 after import; VW fallback lookup for "Virtual Ward Name" in Support sheet
  - `B5_File_Validator.bas` — built and tested
  - `B6_Response_Validator.bas` — updated this session; DT validation checks YYYY-MM-DD 00:00:00.000 format and date range; LS scan uses CStr() on both sides
  - `B6a_DT_Converter.bas` — new this session; TextToColumns with xlDMYFormat to parse DD/MM/YYYY; formats cell as Text (@) before writing string to prevent re-interpretation
  - `B7_Duplicate_Detector.bas` — built this session; Home sheet row comparison on Sub ID + Unique Ref; green colouring; cell-level mismatch detection; logs to Error Log

## Processing sequence (current)
Per file: **Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)**

After all files processed:
1. **B6a_DT_Converter** — parses DD/MM/YYYY text and date serials to YYYY-MM-DD 00:00:00.000 strings
2. **B6_Response_Validator** — orange cell colouring; logs to Error Log sheet
3. **B7_Duplicate_Detector** — green cell colouring; Home sheet row comparison; logs to Error Log sheet
4. **Post to database (A3)** — user-initiated after review

## Home sheet column layout (current)
| Col | Content | Notes |
|---|---|---|
| F | Process? (Yes/No) | FullDataArea col 1; written by B1 on import |
| G | Organisation name | DataArea - 5 |
| H | Submission name | DataArea - 4 |
| I | Org ID | DataArea - 3 |
| J | Sub ID | DataArea - 2 |
| K | CaseCode | DataArea - 1; written back by A3 after successful post |
| L | Unique Ref. | DataArea (col 1) |
| M+ | Question responses | QuestionCols |

## Key constraint confirmed in Session C
**openpyxl drops Form Controls (buttons) on save.** The `.xlsm` must never be passed through the MCP `write_excel` tool once buttons are present.

## Open items
- Drop downs sheet — columns containing numeric list items (e.g. Rockwood scores) must be formatted as Text so XLookup in A3 matches correctly
- Virtual Ward workbook instance not yet built — VW_Data.xlsx has all source data ready to lift across
- DropDownQs named range in VW workbook must cover `'Drop downs'!$A$1:$AI$1` (18 LS questions)
- API `questionType` string for DT — `"date"` used as placeholder; to be confirmed with API team before live use
- VW test files generated but not yet tested end-to-end
