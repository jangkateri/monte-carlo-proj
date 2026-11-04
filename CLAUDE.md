# Monte Carlo demo — working notes

`Project NPV Monte Carlo Simulation.xlsx` is a live Monte Carlo model of a
five-year capital investment. (It was formerly `sample.xlsx`.)
It is driven through the `excel-mcp` server against a visible desktop Excel.

## Prerequisites — required tooling

This project depends on **two** technologies. **Both must be installed and verified
before any other work begins.** If either is missing or broken, run its setup skill
and fix it first — do not open, edit or review workbooks in the meantime, and do not
quietly substitute another library (openpyxl, python-calamine, …) for a missing Excel
MCP server.

| Technology | Setup skill | Healthy when |
|---|---|---|
| Excel MCP Server (`@sbroenne/mcp-server-excel`) | `.claude/skills/excel-mcp-setup` — `/excel-mcp-setup` | `claude mcp get excel-mcp` reports `Status: ✔ Connected`, and the `mcp__excel-mcp__*` tools are present in the session |
| XLRisk add-in (`XLRisk.xlam`) | `.claude/skills/xlrisk-setup` — `/xlrisk-setup` | `XLRisk.xlam` is in `%APPDATA%\Microsoft\AddIns`, an `OPEN*` value under `HKCU\Software\Microsoft\Office\16.0\Excel\Options` names it, and `=RISKNORMAL(100,15)` returns `100` instead of `#NAME?` |

Check both at the start of a session. A registry entry or an `Installed=True` add-in
is **not** proof: XLRisk can be registered yet never initialise its VBA project, and
the Excel MCP server can be registered yet fail to connect. Only the live checks above
count.

### Excel MCP Server

- Install (idempotent — it skips an existing registration):
  ```
  powershell.exe -NoProfile -File "C:\Users\John Ang\monte-carlo-proj\.claude\skills\excel-mcp-setup\scripts\install-excel-mcp.ps1"
  ```
- Requires Windows with an interactive desktop, desktop Excel 2016+, Node.js 18+ and
  `npx`. Manual equivalent: `claude mcp add --scope user excel-mcp -- npx -y @sbroenne/mcp-server-excel@latest`.
- Registration is user scope in `~/.claude.json`. MCP servers load at session start, so
  after installing, Claude Code must be restarted before `mcp__excel-mcp__*` appears.
- A session that reports `Failed to connect — CONNECTION_CLOSED` for `excel-mcp` is a
  setup failure: rerun the skill, don't work around it.

### XLRisk

- Install, pointing at the built add-in (not the `Source` folder):
  ```
  powershell.exe -NoProfile -File "C:\Users\John Ang\monte-carlo-proj\.claude\skills\xlrisk-setup\scripts\install-xlrisk.ps1" -XlamPath "<path to XLRisk.xlam>"
  ```
- Must run under Windows PowerShell 5.1 (`powershell.exe`); PowerShell 7 cannot attach
  to a running Excel.
- Verify outside a simulation, where XLRisk returns the deterministic mean:
  `RiskUniform(0,1)` → `0.5`, `RiskNormal(100,15)` → `100`, `RiskTriang(1,2,3)` → `2`.
  `#NAME?` means the VBA project is not loaded — toggle `Installed` false → true.
  Test cells belong in a scratch workbook closed without saving.
- **On this machine the startup load is dormant every time.** A clean Excel start shows the
  add-in as `Installed=True` yet `Risk*` returns `#NAME?` until it is toggled off and on.
  A full-path `OPEN` value does not change this (tested 2026-11-04; reverted to the relative
  `"XLRisk.xlam"`). Expect one reload per Excel session.

**Excel launched by the MCP runs with `AutomationSecurity = 3 (ForceDisable)`,
so XLRisk never loads there and its ribbon tab is missing.** Toggling
`AddIns(...).Installed` or opening the `.xlam` directly does not fix this
while that flag is set. The only reliable path is a **normal** Excel start
(the user opens the file), where `AutomationSecurity = 1`.

Therefore: prefer **attaching** to an Excel the user opened themselves over
opening a new instance. Verify readiness with
`xl.Evaluate("=RISKNORMAL(0,1)")` — an error variant means the add-in is absent.

## Workbook layout

Three sheets, in this order:

### Overview
Document-style explainer (no gridlines). Prose in `B:D`, rows 2–65.
Section bars at rows 6, 10, 24, 30, 42, 50. Live results table at `B31:C40`
pulls from `Dashboard!E5:E13`. Design-notes table at `B53:D59` maps each
simulation element (inputs, output, iterations) to its cell range. Row heights are set explicitly — wrapped text
needs manual heights, so do not let them be auto-fitted.

### Dashboard
| Range | Contents |
|---|---|
| `A1`, `A2` | Title and subtitle |
| `A4:B11` | Model inputs (yellow). Named ranges below |
| `D4:E14` | Results: trials, mean, std dev, min, 5th pct, median, 95th pct, max, P(NPV>0), base-case NPV |
| `G4:I27` | Histogram. `G6:G26` are **formula-driven bin edges** (start = `Simulation!L2`, step = `Simulation!L1`); `G27` is a dynamic `"> "&TEXT(G26,"#,##0")` label |
| `A15:B21` | "How it works" text |
| `A29` | Dynamic one-line verdict |
| `A31:I37` | Tornado: driver / low / high / NPV at low / NPV at high / downside / upside |
| `K4`, `K22`, `K41` | Chart 1 (histogram), Chart 2 (cumulative), Chart 3 (tornado) |

Named ranges: `InitInv`=B5, `BaseCF`=B6, `GrowthMean`=B7, `GrowthSd`=B8,
`RateMean`=B9, `RateSd`=B10, `Life`=B11.

Tornado layout detail: `B:C` and `E:F` are merged per row (rows 33–36) so the
2.5-wide spacer columns `C` and `F` collapse. Data lives in the merged
anchors only — never write to `C33:C36` or `F33:F36`.

### Simulation
`A1:I2001`. Row 1 headers; rows 2–2001 are formulas (`ROW()`, two
`NORM.INV(RAND(),…)` draws, five cash-flow columns, NPV). Header frozen.

Histogram helpers at `K1:L2`: `L1` = bin step
(`MAX(10,CEILING((MAX−MIN)/21,10))`), `L2` = first edge
(`FLOOR(MIN,L1)+L1`). They re-fit the bins to the current distribution on every
recalculation, so do **not** overwrite `Dashboard!G6:G26` with values.

## Rules when editing

- **NEVER merge a range containing more than one non-empty cell.** Excel
  raises a modal dialog ("Merging cells only keeps the upper-left value")
  that automation cannot dismiss — the session hangs until a human clicks.
  Merge only empty spacer columns, or unmerge → write → re-merge.
- If a dialog appears anyway, ask the user to press **Cancel**, not OK.
  OK discards cells silently.
- Use `overwrite_policy: "allow"` only for cells this work created. Never
  retry a rejected write with `allow` unless the replacement is intended.
- For bulk writes: set calculation to **manual**, write, calculate once,
  restore **automatic**. `RAND()` is volatile and recalculates on every edit.
- A multi-cell write to a merged column is rejected even when every cell is
  an anchor — unmerge first, or write single cells.
- Do not use `window.bring-to-front`, `set-state`, or focus manipulation
  unless asked; Windows blocks activation from background processes and it
  makes things worse.
- Do not save, close, or reopen the workbook without asking. A reopen
  resamples all 2,000 trials and discards the user's view state.
- When a batch fails, inspect state with `file_read list` before retrying.
  Background batches do not roll back.

## Recalculation policy

Every recalculation redraws all 2,000 trials, so numbers never repeat.
If the user wants stable figures, offer to freeze
`Dashboard!E5:E14` and `Simulation!I2:I2001` into values (write with
`overwrite_policy: "allow"` after explicit confirmation).

## Verifying work

Run the `excel-review` skill (`/excel-review`) after any build or modification, and before
handing the workbook over. It runs the checks below plus chart, consistency, layout and — when
the workbook has a VBA project — a read-only review of the macro source.

- `range.get-special-cells` with `cellKind: "Errors"` over
  `Simulation!A1:I2001` and `Dashboard!A1:I40` must return zero areas.
- `range.get-special-cells` with `cellKind: "Constants"` over
  `Simulation!A1:I2001` should return **only** `A1:I1` — anything else means
  a formula row was overwritten by a value (this caught a stray paste in
  `C1`/`C3`).
- Histogram frequencies must sum to the trial count; cumulative must reach 100%.
