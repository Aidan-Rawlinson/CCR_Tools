<!-- Purpose: A snapshot of where the project stands right now -- what works, what is in progress, what is broken. Rewritten by Claude each session. -->

## Status: Exploratory Phase — Alex's Tool Review Incomplete

## Summary
Session 3 documented Alex's tool from the exported `.bas` modules only. The `.xlsm` workbook itself was never inspected — a capability failure that was not flagged at the time. The `Alex_Tool_Reference.md` produced in that session is therefore incomplete: it reflects what the VBA code references, not a full independent inspection of the workbook. That session must be repeated with the `read_excel` tool now available.

The `read_excel` tool has since been built, tested, and is confirmed working. A full structural inspection of the `.xlsm` is now possible.

## What exists
- Project folder structure, Git repository, initial commit on GitHub
- `reference/` folder populated: `Template_Processing_Tool.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, `Guidance(Different tool).pdf`
- `Alex_Tool_Reference.md` in `dynamic_docs/` — **incomplete**: based on `.bas` files only, not a full workbook inspection. To be replaced next session.
- Credentials redacted from `A1_API_SUPPORT.bas` and `.xlsm` (placeholders in place)
- `read_excel` MCP tool built, deployed, and verified working
- No code in `code_base/` yet

## What is in progress
- Nothing — awaiting re-run of Alex's tool documentation session

## What is not started
- Session 3 (redo): Full documentation of Alex's tool using `read_excel` + `.bas` modules
- Session 4: Interpret Alex's tool
- Session 5: Understand the new templates
- Session 6: Design decisions
- Session 7: Test database and API verification
- Session 8: Build Tool 1
- Session 9: Test and iterate Tool 1
- Session 10: Build Tool 2
- Session 11: Test and iterate Tool 2 + final review
