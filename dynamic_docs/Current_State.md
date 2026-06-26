<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Session A Complete — Ready for Session B

## Summary
Session A delivered `B3_Submissions.bas` — a clean, independent module that calls the API and populates the Orgs sheet with all submissions for the configured project in a single operation. Tested successfully against the test database: 25 Managing Frailty submissions pulled first time.

## What exists
- Project folder structure, Git repository, commits on GitHub
- `reference/` folder: Alex's `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete
- Both new questionnaire templates and four SSMS CSVs in `new_questionnaires/`
- Static spec documents: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md` — **still describe old single-file-path flow; not to be relied upon until Session C**
- `dynamic_docs/Colour_Palette.md` — documents Alex's palette and formatting rules
- **`code_base/` contains:**
  - `CCR_Tool_Base.xlsx` — formatted to Alex's palette; Config, Home, Orgs, Drop downs, Lists sheets
  - `managing_frailty_dropdowns.xlsx`
  - `virtual_ward_dropdowns.xlsx`
  - `A1_API_SUPPORT.bas` — VBA-JSON library, UTC utilities, `GetToken()` reading from Config named ranges
  - `A2_API_FUNCTIONS.bas` — all six API functions plus `APICall`/`APIPost` wrappers
  - `A3_API_Calls.bas` — `PostSurveyData` main loop; all five question types (LS, YN, N, TX, DT)
  - `B1_Importer.bas` — will need rework for folder-cycling flow (Session B)
  - `B2_Toggle.bas` — environment toggle handler
  - `B3_Submissions.bas` — **new this session**; `PopulateSubmissions` sub; populates Orgs sheet from API; tested and working

## What is in progress
- Nothing — Session A complete, clean close

## What is not started
- Session B: Folder-cycling importer with org/submission matching support — `B1_Importer.bas` to be reworked
- Session C: Update static spec documents to reflect new flow
- Validation module — `CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator` — parked, separate session
- Tool instances: Home sheet population for Managing Frailty and Virtual Ward
- End-to-end test against test database

## Open items
- Test database access — confirmed working (25 submissions pulled this session)
- API `questionType` string for `DT` — placeholder `"date"` in A3; needs confirming before Virtual Ward test
- Static spec documents need updating to reflect new flow (Session C)
- Matching logic design for Session B — how confident does a match need to be before automatic vs user confirmation? To be agreed at start of Session B
