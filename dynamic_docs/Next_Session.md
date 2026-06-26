<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session A — Orgs Sheet Population

### Context
The planned single-file-path import flow has been replaced with a folder-cycling approach with semi-automated org/submission matching (see Decisions.md). This session focuses on the first prerequisite for that flow: populating the Orgs sheet from the API so the tool knows which organisations are associated with the project.

### Pre-session checklist
- [ ] Test database access confirmed — still needed before any live API calls
- [ ] Confirm API `questionType` string for `DT` questions (Virtual Ward only)
- [ ] User to ensure `CCR_Tool_Base.xlsx` has been saved as `.xlsm` and `.bas` modules imported

### Session A goal
Write and test VBA to call `API_GetSubmissions` and populate the Orgs sheet (Org ID, Org Name, Display string, Submission ID columns) for a given Project ID. This is the foundation the new matching flow depends on.

### Session B goal (after A)
Write and test the folder-cycling importer. The tool cycles over all `.xls*` files in a folder. For each file, it presents the user with enough information to confirm which org and submission it belongs to — or does so automatically where a confident match can be made. Design of the matching logic to be agreed at the start of Session B.

### Session C goal (after B)
Pause build work. Update `Functional_Spec.md`, `Architecture_Design.md`, and `Technical_Spec.md` to reflect the new flow. These are the documents Alex will return to — they must be accurate before the tool instances are built.

### Key open items
- API `questionType` string for `DT` — placeholder `"date"` used in A3
- Validation module (`CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator`) — parked, separate session after core flow is proven
- Static spec documents currently describe the old single-file-path flow — do not rely on them until Session C is complete

### Known risks
- `B1_Importer.bas` was written for the single-file-path approach and will need reworking for the folder-cycling flow — treat it as a starting point, not a finished artefact
- The matching logic design (how confident does a match need to be before it's automatic vs requiring user confirmation?) needs to be agreed before Session B build work begins
