<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session C — Wire up the real importer call

### Context
Session B is complete. `Process_Folder.bas` is written, imported, and tested — all three matching cases working. The stub `ProcessValidFile` confirms a match with a message box. Session C replaces that stub with the real call to `B1_Importer`.

### Pre-session checklist
- [ ] Confirm API `questionType` string for `DT` questions (Virtual Ward only) — not blocking Session C but should be resolved before end-to-end test
- [ ] Orgs sheet should be populated via `B3_Submissions.PopulateSubmissions` before testing the full flow

### Session C goal
Wire `Process_Folder` to `B1_Importer`:

1. **Update `B1_Importer.bas`** — `FileImporter` currently reads `SubmissionFilePath` from the named range and drives its own file path validation. It needs to accept a file path and Submission ID as parameters instead, so `Process_Folder` can call it directly. The named range reference (`SubmissionFilePath` → `SubmissionFolderPath`) also needs updating.

2. **Replace `ProcessValidFile` stub** in `Process_Folder.bas` with a real call to the updated `FileImporter`.

3. **End-to-end test** against test database with a real Managing Frailty submission file.

### Key design question for Session C
`B1_Importer` currently drives its own file path validation (checks file exists, checks sheet exists). When called from `Process_Folder`, the file has already been opened and validated for org/submission matching — some of that validation may be redundant. Decide at the start of Session C whether to:
- Keep validation in `FileImporter` (belt-and-braces, slightly redundant)
- Strip it out and rely on `Process_Folder` having already confirmed the file is accessible

### Still parked
- Validation module (`CaseCodeProcessed`, `QuestionResponseMatcher`, `ResponseValidator`) — separate session after core flow is proven
- Static spec documents (`Functional_Spec.md`, `Architecture_Design.md`, `Technical_Spec.md`) — Session D after Session C is complete
- Tool instances (Home sheet population for Managing Frailty and Virtual Ward) — after spec update
