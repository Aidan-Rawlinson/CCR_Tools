<!-- Purpose: A session-by-session history of what was built and what was decided. The project record. Authored by Claude at session end. -->

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

## Session 3 (redo) — 25 June 2026

**Outcome:** Alex's tool fully documented. `Alex_Tool_Reference.md` rewritten and complete.

- Both `Template_Processing_Tool.xlsm` and `User_Template.xlsx` inspected in full using `read_excel` for the first time
- All five `.bas` modules already documented from Session 3 (original) — not re-read
- Alex's NACEL Submission Importer guidance PDF reviewed (now in Project Files as `GuidanceDifferent_tool.pdf`)
- `Alex_Tool_Reference.md` rewritten to incorporate all findings. Key additions vs previous version:
  - `Lists` sheet documented
  - Home sheet row layout corrected (rows 2–5 were transposed in previous version)
  - Named ranges table expanded with confirmed live values (ProjectID = 40, ServiceID = 146)
  - `SubmissionFolder` dev path issue called out
  - Full User Template section added: sheet structure, section headers, fixed 81-row capacity, two distinct `Drop downs` sheets, service item ID in B1
  - Question count confirmed as 88
  - Known issues consolidated into a single table
  - Guidance document filename corrected to `.docx`
- Notable findings and clarifications:
  - Question number labelling error in Home row 2 (Q51 unlabelled) — confirmed as a data entry error in the workbook
  - BM column holds a database heading stored as a question — no response data expected; blank-skip logic should handle it (verify in Session 4)
  - Template has fixed capacity of 81 rows (ref numbers 101–181)
  - Tool's `Orgs` sheet is a 29-org managed subset; template's `Org list` has 95 orgs (full programme)
  - Two `Drop downs` sheets serve different purposes: template's drives Excel validation (text only); tool's drives import lookup (text + item IDs)
