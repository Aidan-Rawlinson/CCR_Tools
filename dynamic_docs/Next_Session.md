<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session 4 — Interpret Alex's Tool

### Context
`Alex_Tool_Reference.md` is now complete. Session 4 is the interpretation session — the deliberate separation from documentation is a standing project decision. The user brings domain knowledge of the codebase; Claude brings the documented reference. Together, draw out the implications for the new builds.

### Suggested first steps
1. Read `Alex_Tool_Reference.md` in full as the starting point
2. Walk through the tool section by section — not to re-document but to discuss: what is straightforward to replicate, what is project-specific and will need changing, what is unclear
3. Pay particular attention to the Known Issues and Anomalies section — several items need a decision on whether to replicate, fix, or handle differently
4. Agree a clear list of "things we now understand about Alex's tool that have direct implications for the new builds"

### Open questions carried forward
- What `questionType` string does the API expect for free text responses? (`"text"` at ~78% probability, `"string"` at ~21%`) — needs confirming with API team before build
- Two new tools or one generalised tool? (deferred to design session)
- Test database access details — still needed before any build work begins
- Year parameter: `2026` hardcoded in two API URLs — confirm whether dynamic or fixed for this build cycle
- BM column (heading stored as question) — confirm blank-skip logic handles it gracefully; verify against code in this session

### Items from Alex_Tool_Reference.md worth discussing in Session 4
- `SubmissionFolder` reads a folder and processes every `.xls*` file in it — is this the right pattern for new tools, or should it be a single file path?
- Fixed 81-row capacity in the User Template — will new templates follow the same pattern?
- `Drop downs` sheet coverage not guaranteed to be complete — needs cross-referencing against full QID list
- `Org_Id` named range is broken — confirm the code doesn't use it and document accordingly
- Token fetched per API call (no caching) — accept this pattern or improve in new tools?
