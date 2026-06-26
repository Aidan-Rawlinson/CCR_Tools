<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session E Complete — Ready for Session F (Build B6_Response_Validator)

## Summary
Session E completed a significant amount of unplanned but necessary groundwork. The base workbook was fully reconfigured for Managing Frailty (all 38 questions, correct QIDs, types, source row positions, short headers, named ranges extended to AV). The Home sheet gained an "Organisation name" column (G) and a new "Submission name" column (H), with B1 and B4 updated to write both on import. B1's empty-record skip logic was reworked from a unique-ref check to a proper has-data check across all question cells. Test input files were populated with synthetic but valid patient data. Session ordering was corrected: B6_Response_Validator must precede the first live API test.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- `test_inputs/` — four non-broken Managing Frailty test files, each populated with synthetic valid patient data (files 1 & 2: 50 patients, files 3 & 4: 10 patients); four broken files left intentionally broken for B5 testing; `populate_test_files.py` script present
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsm` — fully configured for Managing Frailty: Home rows 2–6 populated (38 questions, QIDs, types, source rows, short headers); named ranges StartCols/QuestionCols/TypeCols/DataArea/FullDataArea all extended to AV; Drop downs sheet populated from managing_frailty_dropdowns.xlsx; Home columns G (Organisation name) and H (Submission name) added; all `.bas` modules imported; buttons in place
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — built and tested
  - `A2_API_FUNCTIONS.bas` — built and tested
  - `A3_API_Calls.bas` — built; all five question types (LS, YN, N, TX, DT)
  - `B1_Importer.bas` — updated this session; signature extended to accept Str_OrgName and Str_SubName; writes org name (DataArea-4), sub name (DataArea-3), sub ID (DataArea-2), unique ref (DataArea) on import; skip logic reworked: checks at least 1 non-blank response across all question cells (StartCols index 2+) rather than unique ref; Lng_ImportedCount tracks records imported; warns if zero records imported across whole file
  - `B2_Toggle.bas` — built and tested
  - `B3_Submissions.bas` — built and tested against test database
  - `B4_Process_Folder.bas` — updated this session; passes Str_OrgName and Str_SubName to B1 at both FileImporter call sites (Case 2 and Case 3)
  - `B5_File_Validator.bas` — built and tested

## Agreed module structure and processing sequence

Processing sequence per file: **Pick (B4) → Validate file (B5) → Match (B4) → Import (B1)**

After all files processed, two further steps run across all imported rows on the Home sheet:
- **Response validation (B6)** — orange cell colouring; validates LS responses against Drop downs sheet, N responses as numeric, TX responses as non-blank text
- **Duplicate detection (B7)** — green cell colouring; API calls to retrieve existing case codes and responses

| Module | Status | Responsibility |
|---|---|---|
| `B1_Importer` | Updated and tested | Pure data transfer; writes org name, sub name, sub ID, unique ref; has-data skip logic |
| `B4_Process_Folder` | Updated and tested | File picker, clear prompt, org/submission matching, calls B5 and B1 |
| `B5_File_Validator` | Built and tested | Structural and content validation of questionnaire files before matching |
| `B6_Response_Validator` | Not started | Response validation: LS (Drop downs lookup), N (numeric check), TX (non-blank); orange cell colouring |
| `B7_Duplicate_Detector` | Not started | Duplicate detection via API; green cell colouring |

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
| F | Build B6_Response_Validator |
| G | First live API test — Managing Frailty (gated on test database access) |
| H | Build B7_Duplicate_Detector |
| I | End-to-end test — Managing Frailty |
| J | Build Managing Frailty tool instance (Home sheet, named ranges, drop downs, import `.bas` files, workbook check) |
| K | Amend for Virtual Ward |

## Open items
- API `questionType` strings for `TX` and `DT` — placeholders in A3; needed before Virtual Ward end-to-end test (Session K); not blocking Sessions F–J
- Test database access — needed before Session G
