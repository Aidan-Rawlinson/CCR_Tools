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
No real data will touch the tool until it has been verified against the test database. Test database access and API endpoint verification is a discrete session (Session 5) that gates all build work.
