<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude at session end. -->

## Handoff for Session F — Build B6_Response_Validator

### Context
Session E is complete. The workbook is fully configured for Managing Frailty, test files are populated, and the import pipeline is working end-to-end. B6 must be built and verified before any live API testing can begin — invalid responses will cause the API post to fail.

### Why B6 must come before the API test
The API import will reject records with invalid response values. B6 is the user's signal to correct data before posting. Running the post against unvalidated data risks creating partial or corrupt case codes against the test database.

### B6 scope — three question types to validate

| Type | Validation rule | Failure response |
|---|---|---|
| `LS` | Response text must match one of the valid options for this question in the Drop downs sheet | Colour cell orange |
| `N` | Cell value must be numeric (IsNumeric check) | Colour cell orange |
| `TX` | No validation needed — blank cells are skipped at post time per existing architecture; non-blank text always valid | No action |
| `YN` | No validation needed — values can only be Yes/No from template drop-down | No action |
| `DT` | Parked — not present in Managing Frailty | Not in scope for this session |

### Design notes for B6

**Inputs:** Home sheet data area — all rows in `FullDataArea` where column F is not blank.

**Named ranges available:**
- `TypeCols` — row 3, K onwards: question type code per column
- `QuestionCols` — row 5, K onwards: QID per column
- `DropDownQs` — Drop downs row 1: QIDs in odd columns (A, C, E...) — used to locate the right column pair
- `DataArea` — data rows starting at K7

**LS lookup logic:** for a given cell, read its QID from `QuestionCols`, find the matching column in `DropDownQs`, then check whether the cell value matches any item in the response text column (even column, rows 3+) for that question. If no match → orange.

**N check:** `IsNumeric(cell.Value)` — if False → orange.

**TX and YN:** skip — no colouring applied.

**Orange colour:** use Alex's convention — interior colour RGB or the standard orange. Check `Colour_Palette.md` — orange is applied dynamically at runtime (not in the static palette) but Alex's code is the reference for the exact colour used.

**Scope:** run across all data rows on Home, not just the most recently imported batch. This matches Alex's pattern and allows B6 to be re-run after manual corrections.

**When called:** B6 is triggered by a separate button on the Home sheet, run after all files have been imported and before the user posts to the database. It is not called automatically by B4.

### Session F goals
1. Review A3_API_Calls.bas to confirm the post flow is ready for Managing Frailty — do this before building B6 so any issues are known
2. Design B6 collaboratively (orange colour value, loop structure, LS lookup approach) before writing code
3. Build and test B6 against the populated test files

### Pre-session checklist
- [ ] Confirm test database access is available (needed for Session G, not F)
- [ ] Confirm API `questionType` strings for `TX` and `DT` with API team (not blocking F)
