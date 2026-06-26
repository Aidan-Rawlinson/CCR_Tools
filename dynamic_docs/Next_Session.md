<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session E — Test Create CCR Records from Managing Frailty Files

### Context
Session D is complete. The full file processing pipeline is wired and tested end-to-end:
- `B1_Importer` accepts parameters, appends rows, has an empty file check
- `B4_Process_Folder` has the clear prompt and calls B1 directly
- `B5_File_Validator` validated and module naming confirmed correct
- Happy path and representative failure cases all confirmed working

Session E is the first live API test — creating real CCR records against the test database from Managing Frailty template files. This is gated on test database access being available.

### Pre-session checklist
- [ ] Confirm test database access is available
- [ ] Confirm API `questionType` strings for `TX` and `DT` with API team — not blocking Session E (Managing Frailty has no DT questions and TX passes through) but needed before Virtual Ward work begins

### Session E goals

**1. Review A3_API_Calls.bas before testing**
Read the current A3 module to confirm the post flow is ready for Managing Frailty — in particular:
- LS question handling (Drop downs lookup)
- YN question handling
- N question handling
- TX question handling (pass-through)
- Correct named range reads

**2. Populate Home sheet metadata for Managing Frailty**
Before the post can run, the Home sheet needs its metadata rows populated for Managing Frailty:
- Row 2: question numbers
- Row 3: question type codes
- Row 4: StartCols values (source row positions in template)
- Row 4: QuestionCols QIDs
- Row 5: column headers
- Named ranges `QuestionCols` and `StartCols` confirmed pointing to correct range

This is build-time configuration — applied directly in Excel by the user.

**3. Copy Managing Frailty Drop downs data into the tool**
Copy the `managing_frailty_dropdowns.xlsx` data into the `Drop downs` sheet of the workbook.

**4. Test the post flow**
- Import a Managing Frailty template file via the full pipeline
- Review imported rows on Home sheet
- Set column F to Yes for selected rows
- Run the post (A3)
- Confirm case codes created and responses posted against test database

### Notes for Session E
- Managing Frailty has no DT questions — TX and LS/YN/N are the types in play
- If any issues arise with A3 named range reads, check against the current Config layout before assuming a code bug
- Test database only — Live toggle must be confirmed as Test before any posting
