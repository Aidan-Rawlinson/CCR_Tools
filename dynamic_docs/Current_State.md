<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session B Complete — Ready for Session C (real importer call)

## Summary
Session B delivered `Process_Folder.bas` — a new module that handles the file picker, org/submission matching decision tree, and end-of-run summary. Tested successfully against four files covering all three matching cases. The stub `ProcessValidFile` confirms a valid match with a message box; replacing it with the real importer call is the next step.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- Static spec documents: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md` — **still describe old single-file-path flow; not to be relied upon until Session C (spec update)**
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsm` — formatted to Alex's palette; Config, Home, Orgs, Drop downs, Lists sheets; `SubmissionFolderPath` named range in place (replaces `SubmissionFilePath`); `DataStart` and `DataMax` named ranges confirmed; Orgs sheet populated with live Managing Frailty submission data
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — VBA-JSON library, UTC utilities, `GetToken()` reading from Config named ranges
  - `A2_API_FUNCTIONS.bas` — all six API functions plus `APICall`/`APIPost` wrappers
  - `A3_API_Calls.bas` — `PostSurveyData` main loop; all five question types (LS, YN, N, TX, DT)
  - `B1_Importer.bas` — single-file importer; still references `SubmissionFilePath` (to be updated when real call is wired in)
  - `B2_Toggle.bas` — environment toggle handler
  - `B3_Submissions.bas` — `PopulateSubmissions` sub; populates Orgs sheet from API; tested and working
  - `Process_Folder.bas` — **new this session**; file picker, org/submission matching, end-of-run summary; tested against four files, all cases working

## What is in progress
- Nothing — Session B complete, clean close

## What is not started
- Session C (importer wire-up): replace `ProcessValidFile` stub in `Process_Folder.bas` with real call to `B1_Importer`; update `B1_Importer.bas` to accept file path and submission ID as parameters rather than reading from named range
- Session D (spec update): update static spec documents to reflect new flow
- Validation module — `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` — parked, separate session
- Tool instances: Home sheet population for Managing Frailty and Virtual Ward
- End-to-end test against test database

## Open items
- API `questionType` string for `DT` — placeholder `"date"` in A3; needs confirming before Virtual Ward test
- Static spec documents need updating to reflect new flow (Session D)
- `B1_Importer.bas` still references `SubmissionFilePath` — to be updated in Session C
