# CCR Tools — Colour Palette & Formatting Rules

Derived from Alex's Template_Processing_Tool.xlsm (reference tool).

---

## Palette

| Token | Hex (ARGB) | Hex (RGB) | Description |
|---|---|---|---|
| `PANEL_BG` | `FFD3DAEE` | `D3DAEE` | Soft blue-grey — panel and control area background |
| `INPUT_CELL` | `FFF9FAFA` | `F9FAFA` | Near-white — user-entry cells within a panel |
| `METADATA_ROW` | `FFFFFF00` | `FFFF00` | Yellow — internal VBA-read rows (hidden from users) |
| `HEADER_DARK` | `FF1F3864` | `1F3864` | Dark navy — sheet and table headers (confirmed from Config/Orgs sheets) |
| `LABEL_ROW` | `FFF2F2F2` | `F2F2F2` | Light grey — label cells adjacent to input cells in tables |

Note: `00000000` (fully transparent) is the default cell state — no fill applied. Used for data rows and standard table content.

---

## Rules

### 1. Control Panel
Apply `PANEL_BG` (`D3DAEE`) to a contiguous block of columns forming the control/config area of a sheet. The panel is purely visual — most cells within it are empty and carry only the background colour. Buttons sit inside the panel without needing their own colour; the panel background is their backdrop.

### 2. Input Cells
Individual cells within a panel where the user enters data get `INPUT_CELL` (`F9FAFA`). The contrast against the panel background is intentionally subtle — just enough lift to signal "type here". Do not use white; keep it within the same cool family as the panel.

### 3. Table Headers
Sheet headers and data table column headers get `HEADER_DARK` (`1F3864`) — the dark navy. White or light text should be used on top of this colour.

### 4. Table Label Cells (Config-style tables)
In two-column label/value tables (e.g. Config sheet), label cells in column A get `LABEL_ROW` (`F2F2F2`) — a neutral light grey that distinguishes the label from the value without competing with the panel or header colours.

### 5. Metadata Rows
Any row that is internal VBA machinery rather than user-facing (question type codes, question IDs, source column positions) gets `METADATA_ROW` (`FFFF00`) across its full column span. These rows are typically hidden from the user after setup.

### 6. Data Rows
No fill. Worksheet default. Colour is only applied dynamically by VBA at runtime:
- **Orange** — invalid response (does not match Drop downs lookup)
- **Green** — likely already imported (case code + responses match database)

---

## Application by Sheet

| Sheet | Panel | Input Cells | Headers | Notes |
|---|---|---|---|---|
| Home | B2:E15 (approx) — `PANEL_BG` | Org selector, submission selector — `INPUT_CELL` | Row 6 column headers — `HEADER_DARK` | Metadata rows 3–5 — `METADATA_ROW` |
| Config | — | B2:B13 value cells — `INPUT_CELL` | A1:B1 — `HEADER_DARK` | A2:A13 label cells — `LABEL_ROW` |
| Orgs | — | — | Row 1 — `HEADER_DARK` | Data rows — no fill |
| Drop downs | — | — | Row 1 (QIDs) — `METADATA_ROW` | Internal sheet; no panel needed |
| Lists | — | — | — | Hidden sheet; no formatting needed |
