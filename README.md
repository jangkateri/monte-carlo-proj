# Monte Carlo modelling workspace

An Excel-based Monte Carlo model of a five-year capital investment, driven through the
[Excel MCP server](https://www.npmjs.com/package/@sbroenne/mcp-server-excel) and, where
needed, the XLRisk add-in. The repository also holds the project's Claude Code skills and
working notes.

## What's here

| Path | Purpose |
|---|---|
| `Project NPV Monte Carlo Simulation.xlsx` | The model. Three sheets: `Overview` (explainer), `Dashboard` (inputs, results, charts), `Simulation` (2,000 formula-driven trials). |
| `Wagesfinal.xlsx` | Chapter 1 exercise from *Microsoft Excel 2016 Data Analysis and Business Modeling* — weekly wages per employee, total hours, average pay. Sheets: `Overview`, `Wages`. |
| `CLAUDE.md` | Working notes for Claude Code: workbook map, editing rules, required tooling. |
| `.claude/skills/` | `excel-review`, `excel-mcp-setup`, `xlrisk-setup`. |

Reference books live in `ebooks/` and are deliberately excluded from version control.

## Prerequisites

Two pieces of software are required. Both must verify before the workbooks are touched.

**1. Excel MCP Server** — needs Windows with an interactive desktop, desktop Excel 2016+ and
Node.js 18+.

```bash
claude mcp add --scope user excel-mcp -- npx -y @sbroenne/mcp-server-excel@latest
claude mcp get excel-mcp        # Status must read "✔ Connected"
```

Or run the guided installer: `/excel-mcp-setup` (script:
`.claude/skills/excel-mcp-setup/scripts/install-excel-mcp.ps1`). Restart Claude Code after
registering — MCP servers load at session start.

**2. XLRisk add-in** — `XLRisk.xlam` is not distributed here; install your own build.

```
/xlrisk-setup     # see .claude/skills/xlrisk-setup
```

Verify with `=RISKNORMAL(100,15)` → `100`, `=RISKUNIFORM(0,1)` → `0.5` (outside a simulation,
XLRisk returns the deterministic mean). `#NAME?` means the VBA project has not initialised —
toggle the add-in's `Installed` property off and on. On this machine that reload is required
once per Excel session.

> **Important:** Excel started by the MCP server runs with `AutomationSecurity = 3`, so XLRisk
> never loads there. For simulation work, open the workbook in Excel yourself and *attach* to
> that instance rather than letting the MCP open a new one.

## Quick start

1. Start desktop Excel and open `Project NPV Monte Carlo Simulation.xlsx`.
2. Start Claude Code in this directory.
3. Ask it to attach to the open workbook (`file_read discover` → `file attach`), then work.
4. Run `/excel-review` before handing anything over.

## Workbook map

### Dashboard

| Range | Contents |
|---|---|
| `A4:B11` | Model inputs (yellow); named ranges below |
| `D4:E14` | Trials, mean, std dev, min, 5th percentile, median, 95th percentile, max, P(NPV>0), base-case NPV |
| `G4:I27` | Histogram; bins are formula-driven from `Simulation!L1:L2` |
| `A29` | One-line verdict |
| `A31:I37` | Tornado sensitivity table |
| `K4`, `K22`, `K41` | Histogram, cumulative and tornado charts |

Named ranges: `InitInv`, `BaseCF`, `GrowthMean`, `GrowthSd`, `RateMean`, `RateSd`, `Life`.

### Simulation

`A1:I2001` — header row plus 2,000 rows of formulas (`ROW()`, two `NORM.INV(RAND(),…)`
draws, five cash-flow columns, NPV). Header frozen.

## Determinism and recalculation

`RAND()` is volatile: every recalculation redraws all 2,000 trials, so figures never repeat.
That is correct for exploring the distribution, and wrong for reporting a number you need to
reproduce. To freeze a run, the totals in `Dashboard!E5:E14` and the NPVs in
`Simulation!I2:I2001` can be replaced with their values — do this deliberately, not by accident.

## Skills

| Skill | Use |
|---|---|
| `/excel-review` | Read-only audit: error cells, formulas overwritten by constants, named ranges, chart series, cross-sheet consistency, layout, and a static review of any VBA modules. |
| `/excel-mcp-setup` | Install and verify the Excel MCP server. |
| `/xlrisk-setup` | Install and verify the XLRisk add-in. |

## Repository layout

```
.
├── .claude/skills/          # excel-review, excel-mcp-setup, xlrisk-setup
├── ebooks/                  # reference PDFs — git-ignored
├── CLAUDE.md                # Claude Code working notes
├── Project NPV Monte Carlo Simulation.xlsx
├── Wagesfinal.xlsx
└── README.md
```

## Version

`v0.9.2` — current release (v0.9.0 was the first tagged release).
