<!-- Purpose: Significant decisions and the reasoning behind them. Kept separate so rationale does not get buried in the Progression_Log. -->

## Decision log

### Excel/VBA as the platform
Inherited from Alex's existing build. Not re-evaluated at this stage — the two new project builds will follow the same approach for consistency and to keep scope manageable.

The fuller reasoning: the preferred approach would be a Python/Streamlit application, but Alex is on paternity leave and it would be disrespectful to rebuild the system mid-year in his absence. He will return to an already-changed landscape (AI-first working practices have accelerated since he left) and the right thing is to deliver within his paradigm. A rebuild conversation can happen when he returns.

### Exploratory phase before planning
Alex is on paternity leave and the existing codebase is not fully understood. An exploratory review of his build is the necessary first step before any planning or development begins.

### Separation of documentation and interpretation
The exploratory phase is split into two discrete sessions: one to document Alex's tool (pure recording, no interpretation), and one to interpret what has been documented (collaborative, with the user's knowledge of the codebase applied). This prevents interpretive assumptions baking in before they have been validated.

### Reference folder as holding area
Alex's .xlsm, exported VBA modules, user instructions, and the two new project templates will be held in a reference/ folder, clearly distinct from code_base/ which is reserved for the actual builds.

### User instructions treated as 95% reliable
Alex's user instructions are a valuable source of intent and user experience context, but are not treated as ground truth. The tool itself is the 100% reliable reference. Anything load-bearing from the instructions is verified against the code and template directly.

### Test database as a non-negotiable gate
No real data will touch the tool until it has been verified against the test database. Test database access and API endpoint verification is a discrete session that gates all build work.

### Credentials to be supplied via input sheet
Credentials will not be hardcoded in new tools. A dedicated input sheet in the Excel workbook will hold username and password, following the same pattern as the `Orgs` sheet holds the toggle. This avoids plaintext credentials in committed code.

### Reference documents to be stored as PDF
Guidance documents with images should be converted to PDF before being dropped into `reference/`. `.docx` files are not directly readable in this environment; PDF is the preferred format for any reference document that needs to be read in future sessions.

### Live/Test toggle pattern carried forward
Alex's Live/Test toggle mechanic — a named range on a hidden sheet, controlled by a shape button — is a proven pattern that solves a real operational need. It will be replicated in the new tools without modification.

### read_excel tool added to MCP server
Direct inspection of `.xlsm` and `.xlsx` files requires openpyxl running on the Windows machine, not in the bash container. A `read_excel` tool was added to the file-reader MCP server to bridge this gap. openpyxl 3.1.5 installed on the Windows machine. This tool is now the standard method for all Excel file inspection in this project.

### Session 3 voided and scheduled for redo
The documentation of Alex's tool in Session 3 was based on `.bas` files only. The `.xlsm` workbook was never directly inspected despite this being the stated intent. The session is treated as incomplete and will be redone in full using `read_excel` combined with the `.bas` modules.
