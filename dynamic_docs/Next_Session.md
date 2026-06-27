<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session H — Build B7_Duplicate_Detector

### Context
Session G is complete. The full pipeline is working end-to-end: import, validate responses (B6), post to database (A3), case code written back, Yes flipped to No. TX confirmed working with placeholder `"text"` questionType. LS lookup confirmed working after fixes to column direction and row start.

### What to do in Session H
Build `B7_Duplicate_Detector.bas`. This module:
- Runs after B6 (called from B4 after B6, or triggered separately — to be agreed)
- Calls the API to retrieve existing case codes for the submission (`GetCaseCodeNote` in A2)
- Compares the unique reference in column L against case code notes from the API
- Colours rows green where a matching case code is found (likely already imported)
- Sets column F to "No" for matched rows (pre-populates the toggle; user retains override)
- Operates on current-run rows only (same first/last row range passed from B4)

### Design questions to resolve at session start
1. Should B7 be called automatically by B4 after B6, or triggered by a separate button?
2. Green colouring — entire row, or just specific columns? Alex's tool coloured response cells green where values matched the database. Confirm desired scope.
3. Should B7 also call `GetCaseCodeResponses` to do a cell-level response comparison (green = match, orange = mismatch on already-imported rows), or is reference-level detection sufficient for now?

### Watch points
- `GetCaseCodeNote` in A2 currently filters to `dataSubmitted = "True"` and `completionStatus = "Completed"` — this means in-progress or failed case codes are not returned. Confirm this is the right filter for duplicate detection.
- The unique reference in column L is "Patient 1", "Patient 2" etc. The case code note field holds the same string (written by A3 at post time). The comparison is a straight string match.
- B7 needs the submission ID per row — available from column J (Sub ID).

### Pre-session checklist
- [ ] Confirm B7 trigger mechanism (auto from B4 vs separate button)
- [ ] Confirm green colouring scope
- [ ] Confirm whether response-level comparison is in scope for this session
