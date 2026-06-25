<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: VBA Codebase Written

## Summary
Session 9 produced all five `.bas` modules. The VBA codebase is complete in first-draft form. `CCR_Tool_Base.xlsx` has also been extended: two new Config rows (`DataStart`, `DataMax`), and the Home sheet now has a full metadata structure with four named ranges (`TypeCols`, `StartCols`, `QuestionCols`, `DataArea`, `FullDataArea`) and a placeholder data table extended to row 19408.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- Static spec documents: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsx` — updated: 11 named ranges on Config, Home metadata structure with 5 named ranges, placeholder table to row 19408
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — VBA-JSON library, UTC utilities, `GetToken()` reading from Config named ranges
  - `A2_API_FUNCTIONS.bas` — all six API functions plus `APICall`/`APIPost` wrappers; `Toggle` and `SubmissionYear` from Config named ranges
  - `A3_API_Calls.bas` — `PostSurveyData` main loop; all five question types (LS, YN, N, TX, DT); reads from `FullDataArea`, `QuestionCols`, `TypeCols` named ranges
  - `B1_Importer.bas` — `FileImporter`; orientation-aware; reads from Config and Home named ranges throughout
  - `B2_Toggle.bas` — environment toggle handler; writes to `Toggle` named range

## What is in progress
- Nothing — awaiting test database access confirmation and Session 10

## What is not started
- Pre-build gate: test database access (must be confirmed before Session 10)
- Validation module — file validation and response validation (parked, separate session)
- `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` — parked, to be treated as a validation module
- Session 10: end-to-end test of VBA against test database
- Session 11: Tool 1 (Managing Frailty) Home sheet built; Drop downs populated; full test
- Session 12: Tool 2 (Virtual Ward) Home sheet built; Drop downs populated; full test; handoff notes for Alex

## Open items
- Test database access — must be confirmed before Session 10 begins
- API `questionType` string for `DT` — needed before end-to-end test
