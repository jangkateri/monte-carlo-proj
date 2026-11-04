---
name: excel-review
description: Audit an open Excel workbook the way a code review audits code - scan every sheet for error cells, cells silently overwritten by values, broken named ranges, dead chart series, inconsistent cross-sheet links and clipped layout, review any VBA modules for defects, then report PASS/FAIL per check. Use when the user asks to validate, review, check, verify or "sanity check" a workbook, after building or modifying one with the Excel MCP, or before handing a workbook over.
---

# Excel workbook review

A read-only audit of the **open** workbook through `excel-mcp`. The goal is not "does it look
right" but "can I prove each part still computes what it claims".

Do not write to the workbook during the review. If a check seems to need a helper cell, skip it
and say so instead. The same rule covers macros: use only the read-only `vba_read` actions
(`status`, `list`, `view`, `read`, `search`, `references`) — never `vba` `import`, `update`,
`replace-procedure`, `delete` or `run`, and never execute a macro to "see what it does".

## 0. Attach, never reopen

1. `file_read discover` — find the already-open workbook.
2. `file attach <workbookId>` — attach; do **not** `file open` if it is already open.
3. Confirm `file_read list` shows `excelState: ready` and `operationPending: false`.
   If it shows `dialogOpen`, stop and ask the user to clear the dialog — do not wait on it.

## 1. Inventory

- `worksheet_read list` — record every sheet; a missing sheet is an immediate FAIL.
- `chart_read list` — record every chart.
- `vba_read status` — does the workbook carry a VBA project, is it password-protected, is project
  access blocked by Trust Center? This decides whether the VBA review in §10 runs or is N/A.
- Read the top-left block of each sheet to learn its layout (see the project `CLAUDE.md` for the
  expected map of `Project NPV Monte Carlo Simulation.xlsx`).

## 2. Error cells — must be zero

For each sheet, over its used range:

```
range_read get-special-cells  cellKind: "Errors"
```

Any returned area is a FAIL. Quote the address and the error (`#VALUE!`, `#REF!`, `#NAME?`) —
`cellErrors` in `range_read get-values` gives the formula and a suggested fix.

## 3. Overwritten formulas — the silent killer

Scan the *formula regions* for constants:

```
range_read get-special-cells  cellKind: "Constants"
```

Build an expected-constants whitelist (headers, labels, input cells). **Any constant outside the
whitelist means a formula was replaced by a typed value** — this is how `Simulation!C1`/`C3` were
caught. Report each one as a FAIL with its address.

## 4. Formula integrity

- `range_read get-formulas` over each formula block. Verify row-to-row relative references advance
  (`$C2` -> `$C3` -> `$C4`); a block that repeats `$C2` everywhere means a bad fill.
- Spot-check that no formula is missing inside a contiguous block (`""` where a formula belongs).
- `range_read validate-formulas` on a sample if the block is generated.

## 5. Named ranges

Read every name listed in `CLAUDE.md` / the model's input block. A name that resolves to `#REF!`
or to the wrong sheet is a FAIL. Confirm each input cell still holds a value, not text.

## 6. Internal consistency (the oracle checks)

These catch a formula that "computes something" but computes the wrong thing:

| Check | Expectation |
|---|---|
| Histogram frequencies `SUM(H6:H27)` | equals the trial count `E5` |
| Cumulative `I27` | equals 100% |
| `E13` P(NPV>0) | between 0 and 1 |
| `E8 <= E9 <= E10 <= E11 <= E12` | min ≤ p5 ≤ median ≤ p95 ≤ max |
| Overview table `B31:C40` vs `Dashboard!E5:E13` | identical values |
| Tornado `H33:I36` | equals `MIN/MAX(E,G)-$E$14` recomputed by hand |
| Base-case NPV `E14` | equals the closed-form NPV of the inputs |

If the model is Monte Carlo, these hold for **any** draw, which makes them usable without a
fixed seed.

## 7. Independent recomputation

Take three sample rows (first, middle, last of the data block). Recompute the output by hand from
the row's inputs and the model's documented formula, then compare to the cell:

> NPV = −InitInv + Σ CFt/(1+r)^t with CFt = BaseCF × (1+g)^t

A mismatch beyond rounding is a FAIL. This is the only check that catches a wrong-but-non-erroring
formula (sign flipped, wrong period, wrong rate).

## 8. Chart integrity

`chart_read read <name>` for each chart:

- every series must report a `sourceRange` of `SERIES(...)` pointing at cells — a series holding a
  fixed array instead of a range is FAIL (it will never update)
- categories non-empty and matching the intended labels
- title present; legend appropriate for the series count

## 9. Layout

`screenshot capture-sheet` per sheet (or `capture` a range). Look for:
clipped headers, text overlapping the next cell, charts covering the table, content below the
visible window, empty merged cells inside a table.

Note: screenshots require an interactive desktop; if unavailable, say the check was skipped.

## 10. VBA code review (only when the workbook has macros)

Run this whenever §1 found a VBA project — a `.xlsm`/`.xlsb` whose behaviour depends on macros —
not only when a sheet looks wrong. A workbook can pass every cell check and still be broken
because a macro overwrites the data on open.

Gather the source, read-only:

```
vba_read status                      # access, password protection, execution mode
vba_read list                        # every module and procedure
vba_read view   moduleName: "Module1"
vba_read read   procedureName: "..."   # or startLine / lineCount for a long module
vba_read search searchText: "ActiveSheet"   # literal text, wildcards not supported
vba_read references                  # library references; flags missing ones
```

If `status` reports access blocked, stop and report **VBA review SKIPPED — project access
blocked**. Do not change Trust Center settings on the user's behalf. A password-protected project
cannot be reviewed: say so and ask for the password or an exported copy.

If the user enables *Trust access to the VBA project object model* themselves, warn them that
Excel reads `AccessVBOM` only at startup and writes it back on exit: close every Excel window
first, change the setting, then reopen — editing it while Excel runs gets silently reverted.

The six read actions are verified against a live `.xlsm` (2026-11-04): `status` returns
`projectAccess`, `protection`, `mode`; `list` returns modules, procedures and start lines;
`view` returns module source; `read` returns a procedure body plus `sourceHash`; `search` returns
matches with module, line, column and excerpt; `references` returns libraries with `isBroken`.

Review the source against this list. Report by severity; a finding is only worth listing if you
can name the failure it causes.

| # | Check | Failure it causes |
|---|---|---|
| 1 | Unqualified `Range`/`Cells`/`Rows`/`Columns`, `ActiveSheet` | acts on the wrong sheet when run from elsewhere — silent data corruption |
| 2 | Hardcoded row/column numbers, paths, filenames | breaks as soon as data grows or the file moves |
| 3 | `On Error Resume Next` with no `On Error GoTo 0` | errors swallowed; the macro reports success while doing nothing |
| 4 | `ScreenUpdating`/`Calculation`/`EnableEvents`/`DisplayAlerts` not restored, especially on the error path | Excel left in manual calc, events dead, window frozen |
| 5 | `.Select`/`.Activate`/`Selection`/`ActiveCell` | slow, and fails whenever focus differs; hides which range is meant |
| 6 | Missing `Option Explicit`, undeclared or Variant variables | typos compile silently into new variables |
| 7 | Cell-by-cell loops instead of array reads/writes | minutes instead of seconds on large ranges |
| 8 | Unbounded `Do`/`While`, or recursion with no exit | Excel hangs |
| 9 | `references` shows a missing library | compile error, or a late-bound call failing at runtime |
| 10 | `Workbook_Open` or event handlers with side effects; `Worksheet_Change` with no re-entry guard | data overwritten on open, or the handler retriggers itself |
| 11 | Dead procedures, commented-out blocks, duplicated logic | drift between what the workbook claims and what actually runs |

Also flag anything that writes over the model's inputs, unprotects sheets, shells out, or stores
a credential in source.

Do **not** judge a macro by running it. This is a static review: MCP-run Excel has
`AutomationSecurity = 3`, so macros cannot execute there anyway, and running one against the
user's data is a write, not a review.

Result: PASS if nothing above applies; FAIL with one row per finding (module, procedure, line,
defect, consequence); **N/A** when `vba_read status` shows no project.

## 11. Volatility

Count volatile functions (`RAND`, `NOW`, `TODAY`, `OFFSET`, `INDIRECT`). If any exist, state that
every recalculation changes the numbers, and offer to freeze values before the workbook is used
for reporting.

## Report

Return a compact table and nothing else unless asked:

| # | Check | Scope | Result | Evidence |
|---|---|---|---|---|
| 2 | Error cells | Simulation A1:I2001 | PASS | 0 areas |
| 3 | Overwritten formulas | Simulation A1:I2001 | FAIL | C3 holds text, expected formula |
| 10 | VBA code review | Module1 · Refresh | FAIL | `Range("A1")` unqualified — writes to whichever sheet is active |

For a workbook with no VBA project, report check 10 as N/A with the evidence `vba_read status:
no project` rather than omitting the row.

Finish with one line: **Review passed** (all checks PASS) or **Review failed — N issues**, followed
by the ordered fix list. Never say "looks good" without the table.
