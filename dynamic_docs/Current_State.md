<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Design Complete — SSMS Data Ready

## Summary
Session 6 completed the data preparation phase. QIDs and list item IDs have been retrieved from SSMS, cross-referenced against both questionnaire templates, cleaned, and saved as CSVs. The project is ready to build the Drop downs sheets and base workbook structure in the next session.

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
- No code in `code_base/` yet

## What is in progress
- Nothing — awaiting Session 7 (Drop downs sheets + base workbook structure)

## What is not started
- Pre-build gate: test database access (must be confirmed before Session 8)
- Session 7: Drop downs sheets built from CSVs; base workbook structure created
- Session 8: API layer VBA (`A1_API_SUPPORT`, `A2_API_FUNCTIONS`)
- Session 9: Importer VBA (`B1_Importer`)
- Session 10: Orchestration and UI VBA (`A3_API_Calls`, `B2_Toggle`); end-to-end test
- Session 11: Tool 1 (Managing Frailty) configured and tested
- Session 12: Tool 2 (Virtual Ward) configured and tested; handoff notes for Alex

## Open items
- ⚠️ **BLOCKING (Session 8):** API date format for `DT` question type — must be confirmed with API team before importer is written
- ⚠️ **BLOCKING (Session 8):** API `questionType` string for `DT` questions — must be confirmed before Session 8
- Test database access — needed before Session 8
- API `questionType` string for free text (`TX`) — needed before Session 10
