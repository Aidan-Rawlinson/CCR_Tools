<!-- Purpose: Claude's handoff note -- what to pick up, open questions, and suggested first steps for the next session. Written by Claude each session. -->

## Handoff for Session I — Build Virtual Ward Workbook Instance

### Context
Session H is complete. B7 is built and wired. B6a handles DT conversion. B4 has the VW fallback lookup. All module updates for LS numeric matching and DT are in place. VW_Data.xlsx contains the correct Config, Home rows 2–6, and Drop downs content ready to lift across.

### What to do in Session I

1. **Format Drop downs numeric columns as Text** — before lifting VW_Data.xlsx content across, ensure any Drop downs columns containing numeric list items (Rockwood scores: column S in VW Drop downs) are formatted as Text in the target workbook. This ensures XLookup in A3 matches correctly.

2. **Build the Virtual Ward workbook instance** — create `CCR_Tool_VirtualWard.xlsm`:
   - Start from CCR_Tool_Base.xlsm as the base (copy it)
   - Lift Config data from VW_Data.xlsx (B2, B13, B14, B15 are the key differences from MF)
   - Lift Home rows 2–6 from VW_Data.xlsx (question numbers, types, StartCols row numbers, QIDs, headers)
   - Lift Drop downs content from VW_Data.xlsx
   - Update all named ranges: DataArea, FullDataArea, QuestionCols, StartCols, TypeCols to cover correct column range for 33 VW questions (L through AR)
   - Update DropDownQs named range to `'Drop downs'!$A$1:$AI$1`
   - Import all .bas modules
   - Reinstate buttons

3. **End-to-end test — Virtual Ward** — use the three VW test files in test_inputs/ against the test database

### Watch points
- VW template orientation is Columns, DataStart = E — same as MF
- VW has 33 questions (L through AR on Home sheet), one DT question (62624, Referral date)
- SpotChecks in Config are VW-specific — confirm they pass B5 validation against the test files
- Essex Partnership has two submissions in the Orgs sheet — picker will present multiple-match case; confirm correct submission selected
- B4 VW fallback (Support!A6 = "Virtual Ward Name") is already in the .bas file

### Pre-session checklist
- [ ] Confirm DropDownQs range column count for VW (should be A1:AI1 — 18 LS questions × 2 cols)
- [ ] Confirm all named ranges updated correctly in VW workbook before importing modules
- [ ] Confirm DT questionType string with API team before live posting
