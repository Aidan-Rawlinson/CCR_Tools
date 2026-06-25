<!-- Purpose: A session-by-session history of what was built and what was decided. The project record. Authored by Claude at session level, not micro-decision level. -->

## Session 1 — 24 June 2026

**Outcome:** Project setup and briefing only. No code written.

- Git repository initialised on the CCR_Tools project folder and initial commit pushed to GitHub
- High-level brief established: the tool processes bulk uploads of Clinical Case Note Review data from Excel templates into TBN's online data collection system
- Alex's existing Excel/VBA build (for a different project) identified as the reference point
- Two new projects identified as the target scope
- Agreed approach: start with an exploratory phase to understand Alex's build before planning the adapted versions

## Session 2 — 24 June 2026

**Outcome:** Planning phase completed. No code written.

- Established working approach and division of labour: user is a competent VBA programmer, primary risk is data flow and user experience understanding rather than coding
- Confirmed API-to-database architecture: Alex's tool sends data via API to TBN's master database(s)
- Confirmed test database requirement as a non-negotiable gate before any real data is used
- Agreed to stay in Excel/VBA and follow Alex's approach faithfully — rationale documented in Decisions
- Established that Alex's user instructions exist and should be read first in the exploratory phase, weighted at ~95% reliable with the tool itself as ground truth
- Confirmed capabilities: Claude can read .xlsm structure via openpyxl (available), read exported .bas modules as plain text, and read PDFs directly — full sight of the tool is achievable
- Agreed reference folder structure for Alex's files and new project templates
- Agreed to separate documentation and interpretation into discrete sessions
- Produced a 9-session plan (see Next_Session)

## Session 3 — 25 June 2026

**Outcome:** Partial documentation only. Session voided — to be redone.

- Reference folder populated by user: `.xlsm`, `User_Template.xlsx`, five exported `.bas` modules, guidance PDF (different project clone)
- All five VBA modules read and documented
- `Alex_Tool_Reference.md` produced — but based on `.bas` files only; the `.xlsm` workbook was never directly inspected
- **Failure:** Claude stated a capability (reading `.xlsm` structure via openpyxl) in Session 2 that was not delivered in Session 3. The workbook was never passed through openpyxl. Structural knowledge of the workbook was inferred from VBA code references rather than direct inspection. This was not flagged at the time.
- The `Alex_Tool_Reference.md` is therefore incomplete and cannot be treated as a reliable foundation for build decisions
- Guidance document reviewed and discrepancies recorded — this element of the session remains valid
- Credentials redacted from `.bas` and `.xlsm` — this element of the session remains valid

## Session 4 — 25 June 2026

**Outcome:** Capability gap identified and remediated. Session plan rolled back.

- Identified that `read_excel` MCP tool was needed to properly inspect `.xlsm` and `.xlsx` files
- Built and deployed `read_excel` tool to the MCP server (openpyxl installed on Windows machine, server updated)
- Tested `read_excel` against `Template_Processing_Tool.xlsm` — confirmed working, full structural data returned
- First full inspection of `.xlsm` revealed several details absent from `Alex_Tool_Reference.md`: `Lists` sheet, full `Orgs` sheet org data, `Home` row 2 structure (question numbers), `Org_Id` broken named range
- Two new project templates identified (Managing Frailty, Virtual Ward) — inspection deferred pending session reset
- Two architectural scope changes identified for the new builds:
  - Transposed template layout (questions in rows, patients in columns) vs Alex's layout (questions in columns, patients in rows)
  - Four response types to support (`response_list_id`, `response_txt`, `response_yn`, `response_num`) vs Alex's three (`LS`, `YN`, `N`) — `TX` (free text) type to be confirmed with API team
- Decision taken: roll back to redo Session 3 properly before proceeding
