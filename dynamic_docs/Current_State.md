<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Base Workbook Built

## Summary
Session 8 produced `CCR_Tool_Base.xlsx` — the base workbook structure for all three tool instances. The `environment/` folder was also cleaned up: scratch files removed, Drop downs sheets moved to `code_base/`, and the folder boundary between build artefacts and environment infrastructure is now clearly established.

## What exists
- Project folder structure, Git repository, initial commit on GitHub
- `reference/` folder: `Template_Processing_Tool.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete, based on full workbook inspection plus `.bas` modules
- Both new questionnaire templates in `new_questionnaires/`: `NHSBN Managing Frailty...xlsx` and `NHSBN Virtual Ward...xlsx`
- Four SSMS CSVs in `new_questionnaires/`:
  - `managing_frailty_question_ids.csv` — 42 questions, cleaned IDs
  - `managing_frailty_list_item_ids.csv` — list items for all LS questions
  - `virtual_ward_question_ids.csv` — 33 questions, cleaned IDs
  - `virtual_ward_list_item_ids.csv` — list items for all LS questions
- Solution architecture agreed: one generalised VBA codebase, three workbook instances (Base, Managing Frailty, Virtual Ward)
- Build plan agreed: 7 build sessions (Sessions 6–11 revised to Sessions 6–13)
- `write_excel` MCP tool built, deployed to live `server.py`, and verified working
- Static spec documents populated: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`
- All five question types fully designed: `LS`, `YN`, `N`, `TX`, `DT`
- **`code_base/` now contains three build artefacts:**
  - `CCR_Tool_Base.xlsx` — base workbook; all five sheets, Config layout, all 9 named ranges, data validation drop-downs for Toggle and Orientation
  - `managing_frailty_dropdowns.xlsx` — 10 LS questions, columns A–T
  - `virtual_ward_dropdowns.xlsx` — 18 LS questions, columns A–AJ
- **`environment/` contains only:** `git_push.py`, `git_revert.py`, `git_log.txt`

## What is in progress
- Nothing — awaiting Session 9 (API layer VBA)

## What is not started
- Pre-build gate: test database access (must be confirmed before Session 9)
- Session 9: API layer VBA (`A1_API_SUPPORT`, `A2_API_FUNCTIONS`)
- Session 10: Importer VBA (`B1_Importer`)
- Session 11: Orchestration and UI VBA (`A3_API_Calls`, `B2_Toggle`); end-to-end test
- Session 12: Tool 1 (Managing Frailty) configured and tested
- Session 13: Tool 2 (Virtual Ward) configured and tested; handoff notes for Alex

## Open items
- Test database access — must be confirmed before Session 9 begins
- API `questionType` string for `DT` — needed before Session 11
