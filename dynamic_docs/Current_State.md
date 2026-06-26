<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Flow Redesign — Pre-Build Pause

## Summary
Session 10 has produced three things: MCP server improvements (colour reading, in-place editing, font colour), a colour palette document, and formatting applied to CCR_Tool_Base.xlsx. More significantly, a review of Alex's code against the originally planned user flow has identified that the flow needs to change. The planned single-file-path approach is being replaced with a folder-cycling approach with semi-automated org/submission matching. This requires a design update before build resumes.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- Static spec documents: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md` — **to be updated to reflect new flow**
- `dynamic_docs/Colour_Palette.md` — new this session; documents Alex's palette and formatting rules
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsx` — updated this session with full palette formatting (panel, headers, input cells, font colours)
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — VBA-JSON library, UTC utilities, `GetToken()` reading from Config named ranges
  - `A2_API_FUNCTIONS.bas` — all six API functions plus `APICall`/`APIPost` wrappers; `Toggle` and `SubmissionYear` from Config named ranges
  - `A3_API_Calls.bas` — `PostSurveyData` main loop; all five question types (LS, YN, N, TX, DT); reads from `FullDataArea`, `QuestionCols`, `TypeCols` named ranges
  - `B1_Importer.bas` — `FileImporter`; orientation-aware; reads from Config and Home named ranges throughout — **will need rework for new flow**
  - `B2_Toggle.bas` — environment toggle handler; writes to `Toggle` named range
- MCP server updated: colour reading in `read_excel`, in-place editing in `write_excel`, `set_font_colour` op added, alpha opacity fix for `set_background_colour`

## What is in progress
- Flow redesign: the single-file-path import approach is being replaced with folder-cycling with semi-automated org/submission matching. VBA and design docs to be updated.

## What is not started
- Session A: Write & test VBA to populate Orgs sheet from API
- Session B: Write & test folder-cycling importer with org/submission matching support
- Session C: Pause and write documentation for new flow (update Functional_Spec, Architecture_Design, Technical_Spec)
- Validation module — `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` — parked, separate session
- Tool instances: Home sheet population for Managing Frailty and Virtual Ward
- End-to-end test against test database

## Open items
- Test database access — still needed before end-to-end test
- API `questionType` string for `DT` — placeholder `"date"` in A3; needs confirming before Virtual Ward test
- Static spec documents need updating to reflect new flow
