<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Design Complete — Ready to Build

## Summary
Session 5 (this session) completed the interpretation and design phase. Both new questionnaire templates have been inspected, the solution architecture is agreed, the build plan is confirmed, and the `write_excel` MCP tool has been built and verified. The project is ready to enter the build phase.

## What exists
- Project folder structure, Git repository, initial commit on GitHub
- `reference/` folder: `Template_Processing_Tool.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).docx`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — complete, based on full workbook inspection plus `.bas` modules
- Both new questionnaire templates inspected: `NHSBN Managing Frailty...xlsx` and `NHSBN Virtual Ward...xlsx`
- Solution architecture agreed: one generalised VBA codebase, three workbook instances (Base, Managing Frailty, Virtual Ward)
- Build plan agreed: 7 build sessions (Sessions 5–11)
- `write_excel` MCP tool built, deployed to live `server.py`, and verified working
- Static spec documents populated: `Project_Brief.md`, `Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`
- Credentials approach agreed: Config sheet input, not hardcoded
- No code in `code_base/` yet

## What is in progress
- Nothing — awaiting Session 6 (SSMS data + Drop downs sheet build)

## What is not started
- Pre-build gate: test database access (must be confirmed before Session 7)
- Session 6: SSMS CSVs supplied; QID and list item ID lookup sheets built
- Session 7: API layer VBA (`A1_API_SUPPORT`, `A2_API_FUNCTIONS`)
- Session 8: Importer VBA (`B1_Importer`)
- Session 9: Orchestration and UI VBA (`A3_API_Calls`, `B2_Toggle`); end-to-end test
- Session 10: Tool 1 (Managing Frailty) configured and tested
- Session 11: Tool 2 (Virtual Ward) configured and tested; handoff notes for Alex

## Open items
- Test database access — needed before Session 7
- QID CSVs for both projects — needed for Session 6
- List item ID CSVs for both projects — needed for Session 6
- API `questionType` string for free text (`TX`) — needed before Session 9
