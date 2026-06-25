<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session 4 (next working session)

### Context
Alex's tool is fully documented and understood. The session plan has shifted slightly — Sessions 1 and 2 of the original plan (document and interpret) have been effectively merged into this single session. The next step is to bring in the two new project templates and map them against Alex's.

### Pre-session action required from user
Drop the two new project templates (Excel files) into the `reference/` folder before the session begins.

### Suggested first steps
1. Read the two new templates — sheet structure, data layout, question columns
2. Map against Alex's `User_Template.xlsx` — what is common, what differs, what is missing
3. Identify project-specific values that will need to change: sheet name (currently `Bed based CCR`), column ranges, row offsets

### Session plan (revised)

**Session 4 — Understand the New Templates** *(next)*
Bring in the two new project templates. Map their structure against Alex's. Identify what is common, what differs, what is missing.

**Session 5 — Design Decisions**
Based on template mapping, make the key calls: two separate tools or one generalised tool, how closely to follow Alex's architecture, how to handle any structural differences. Produce a design document.

**Session 6 — Test Database & API Verification**
Establish that the test database is accessible and the API endpoint works. Verify authentication and payload structure. Gate before any build work.

**Session 7 — Build Tool 1**
Develop the first new tool against the test database.

**Session 8 — Test & Iterate Tool 1**
Realistic data testing. Work through issues, edge cases, validation gaps. Tool 1 reaches finished state.

**Session 9 — Build Tool 2**
Develop the second tool. Pattern established from Tool 1 — this session should be faster.

**Session 10 — Test & Iterate Tool 2 + Final Review**
Same as Session 8 for Tool 2. Close with a review of both tools together — consistency, documentation, handoff notes for Alex.

### Open questions
- What are the two new project templates and how do they differ structurally from Alex's?
- Two separate tools or one generalised tool? (To be resolved in Session 5)
- Test database access details and who controls it
- Any constraints on the Excel/VBA environment (macro security settings, Excel version in use by clinicians)
- Year parameter: `2026` is hardcoded in two API URLs — confirm whether this needs to be dynamic or can remain hardcoded for the current build cycle
