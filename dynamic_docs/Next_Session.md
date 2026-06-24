<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session 3 (first working session)

### Pre-session action required from user
Before Session 3 begins, drop the following into a reference/ folder in the project:
- Alex's .xlsm file
- Exported VBA modules (.bas files) from Alex's tool
- Alex's user instructions (PDF)
- The two new project templates (Excel files)

### Session plan (9 sessions)

**Session 1 — Document Alex's Tool**
Read user instructions, .xlsm structure, and VBA modules. Pure recording — no interpretation. Output is a set of structured reference documents in the datastore.

**Session 2 — Interpret Alex's Tool**
Collaborative. Talk through the documentation produced in Session 1. Agree what we are seeing — data flow, coupling points, validation logic, API call structure. Document conclusions.

**Session 3 — Understand the New Templates**
Bring in the two new project templates. Map their structure against Alex's. Identify what is common, what differs, what is missing.

**Session 4 — Design Decisions**
Based on Sessions 1-3, make the key calls: two separate tools or one generalised tool, how closely to follow Alex's architecture, how to handle validation differences. Produce a design document.

**Session 5 — Test Database & API Verification**
Establish that the test database is accessible and the API endpoint works. Verify authentication and payload structure. Gate before any build work.

**Session 6 — Build Tool 1**
Develop the first new tool against the test database.

**Session 7 — Test & Iterate Tool 1**
Realistic data testing. Work through issues, edge cases, validation gaps. Tool 1 reaches finished state.

**Session 8 — Build Tool 2**
Develop the second tool. Pattern established from Tool 1 — this session should be faster.

**Session 9 — Test & Iterate Tool 2 + Final Review**
Same as Session 7 for Tool 2. Close with a review of both tools together — consistency, documentation, handoff notes for Alex.

### Open questions
- How different are the two new project templates from Alex's?
- Two separate tools or one generalised tool? (To be resolved in Session 4)
- Full shape of the API — endpoints, authentication, payload structure
- Test database access details and who controls it
- Any constraints on the Excel/VBA environment (macro security settings, Excel version in use by clinicians)
