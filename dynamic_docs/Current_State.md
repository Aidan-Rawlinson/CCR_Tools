<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Drop Downs Sheets Complete

## Summary
Session 7 completed the Drop downs sheets for both tool instances. Pre-session design work resolved all outstanding blocking items on question type handling. The base workbook structure is the next step.

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
- Build plan agreed: 7 build sessions (Sessions 6–11)
- `write_excel` MCP tool built, deployed to live `server.py`, and verified working
- Static spec documents populated: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`
- All five question types fully designed: `LS`, `YN`, `N`, `TX`, `DT`
- **Drop downs sheets built and formatted:**
  - `environment/managing_frailty_dropdowns.xlsx` — 10 LS questions, columns A–T
  - `environment/virtual_ward_dropdowns.xlsx` — 18 LS questions, columns A–AJ
  - Both formatted: navy QID row, grey-blue label row, alternating blue/white column pairs in data rows
- No code in `code_base/` yet

## What is in progress
- Nothing — awaiting Session 8 (base workbook structure)

## What is not started
- Pre-build gate: test database access (must be confirmed before Session 9)
- Session 8: Base workbook structure (`CCR_Tool_Base.xlsx`) — all sheets, Config layout, named ranges, data validation
- Session 9: API layer VBA (`A1_API_SUPPORT`, `A2_API_FUNCTIONS`)
- Session 10: Importer VBA (`B1_Importer`)
- Session 11: Orchestration and UI VBA (`A3_API_Calls`, `B2_Toggle`); end-to-end test
- Session 12: Tool 1 (Managing Frailty) configured and tested
- Session 13: Tool 2 (Virtual Ward) configured and tested; handoff notes for Alex

## Open items
- Test database access — needed before Session 9 (only remaining pre-build gate)
- API `questionType` string for `DT` — needed before Session 10
