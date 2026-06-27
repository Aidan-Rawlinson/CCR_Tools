<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session G — First Live API Test

### Context
Session F is complete. B6 is built, tested, and passing cleanly on all five question types. The full pipeline — pick, validate, match, import, validate responses — is working end-to-end. The tool is ready for live API testing against the test database.

### Gate
Session G is blocked on test database access. This must be confirmed available before starting.

### What to do in Session G
1. Confirm credentials are entered in Config (APIUsername / APIPassword)
2. Confirm Toggle is set to Test
3. Run Populate Submissions — verify Orgs sheet populates correctly for Managing Frailty (Project 35)
4. Process one test file — verify import, B6 validation, and Home sheet layout all look correct
5. Post one row — step through A3 manually if needed; verify case code written back to column J
6. Check the test database directly via SSMS to confirm the record landed correctly

### Watch points
- A3 reads `Rng_DataRows` as column 5 of `FullDataArea` — this is column J (CaseCode), not column K (Unique Ref). Double-check the offset arithmetic lands on column F for the Yes/No toggle before posting.
- The `Lng_OrgId` split on `"-"` in A3 (`Split(Rng_Cell.Offset(0, -3).Value, "-")(0)`) — column I holds the Sub ID (a plain integer), not an org-dashed string. Verify this doesn't error on clean data; if it does, a small fix to read org ID differently will be needed.
- TX and DT questionType strings in A3 are placeholders (`"text"` and `"date"`). These may or may not be correct — if the API rejects them, note the error response and flag for the API team.

### Pre-session checklist
- [ ] Test database access confirmed
- [ ] Credentials available for Config entry
- [ ] At least one clean test file ready (files 1–4 in test_inputs/ are pre-populated)
