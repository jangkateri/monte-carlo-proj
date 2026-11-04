---
name: xlrisk-setup
description: Install the XLRisk Monte Carlo add-in into desktop Excel on Windows for the first time - copy the .xlam to the per-user add-ins folder, register it in the Excel Options registry key so it loads on every start, force a reload, and verify that Risk* worksheet functions evaluate. Use when the user asks to install or set up XLRisk, or when XLRisk functions return #NAME?.
---

# XLRisk Setup (Windows desktop Excel)

Installs `XLRisk.xlam` as a per-user Excel add-in and proves the `Risk*` UDFs are live.
Install is not complete until a test formula returns a value instead of `#NAME?`.

## Decision flow

1. **Find the add-in build.** You need `XLRisk.xlam`, not the `Source` folder (that is VBA
   source that must be built into an `.xlam`). Likely locations:
   - `%USERPROFILE%\monte-carlo-sim\XLRisk\XLRisk.xlam` (this machine's repo)
   - `%APPDATA%\Microsoft\AddIns\XLRisk.xlam` (an existing install)
   - otherwise ask the user
2. **Check for an existing install before copying:**
   - `Test-Path "$env:APPDATA\Microsoft\AddIns\XLRisk.xlam"`
   - `HKCU:\Software\Microsoft\Office\16.0\Excel\Options` -> any `OPEN*` value naming `XLRisk.xlam`
   - Excel also lists `.xlam` files found in the AddIns folder as `Installed=False`; that is
     "available", not installed.
3. **Run the install script with Windows PowerShell 5.1** (PowerShell 7 cannot attach to a
   running Excel; `Marshal.GetActiveObject` does not exist there):
   ```
   powershell.exe -NoProfile -File "C:\Users\John Ang\monte-carlo-proj\.claude\skills\xlrisk-setup\scripts\install-xlrisk.ps1" -XlamPath "<path to XLRisk.xlam>"
   ```
   Add `-NoVerify` only if the user explicitly does not want Excel touched.
4. **Report** the installed path, the `OPEN` registry value, and the verification table.
   Exit code 0 = working; 1 = registration or evaluation failed (see troubleshooting).

## What the script does

1. Probes the running Excel (or starts a hidden one) to learn the version and whether an
   `XLRisk*` add-in entry already exists, then quits the probe if it created it.
2. Copies the `.xlam` into `%APPDATA%\Microsoft\AddIns` (only when the hash differs),
   reusing an existing registration path when one exists.
3. Registers autoload by writing `"<path>"` to the first free `OPEN` / `OPEN1` / ... value
   under `HKCU:\Software\Microsoft\Office\<ver>\Excel\Options` - the same key Excel writes.
   It does **not** call `Application.AddIns.Add`, which fails under automation with
   "Unable to get the Add property of the AddIns class".
4. Reloads the add-in in the user's already-open Excel (`Installed` false -> true) without
   closing any workbooks. If the running Excel predates the registration, it warns that a
   restart is needed instead.
5. Verifies in a **fresh** hidden Excel instance started after registration, reloading the
   add-in there too, then quits that instance.

## Why the reload is mandatory

Excel can load an add-in at startup yet never initialise its VBA project: the entry reports
`Installed=True` and every `Risk*` call returns `#NAME?`. Setting `Installed = $false` then
`$true` forces a real load. This happens even in instances Excel starts fresh, so always
reload before verifying and before telling the user it works.

## Verification facts

Outside a simulation XLRisk returns the **deterministic mean** (`ProduceRandomSample = False`
in `RiskFunctions.bas`). This makes install verification exact:

| Formula | Expected |
|---|---|
| `=RiskUniform(0,1)` | `0.5` |
| `=RiskNormal(100,15)` | `100` |
| `=RiskTriang(1,2,3)` | `2` |

Random values appear only while a simulation runs. `#NAME?` means the VBA is not loaded.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `Unable to get the Add property of the AddIns class` | `AddIns.Add` is blocked in automation-started Excel | Use the registry `OPEN` value (the script does) |
| All `Risk*` return `#NAME?`, entry shows `Installed=True` | Dormant startup load | Toggle `Installed` false -> true, recalculate |
| `#NAME?` persists after reload | `Application.AutomationSecurity = 3` (ForceDisable) on an automation-started Excel, e.g. the Excel MCP server | Use a normally started Excel, or set `AutomationSecurity = 1` before reloading. Excel MCP sessions force macros off, so XLRisk never works there |
| Add-in unchecks itself | Trust Center "Disable all Application Add-ins", or `VBAWarnings` 3/4 policy | Inspect `HKCU:` and `HKLM:\Software\Policies\Microsoft\Office\16.0\Excel\Security`; enabling is the user's Trust Center action |
| Copy fails, add-in will not load | File locked by Excel, or blocked by Mark of the Web | Unload first (the script does); check `Get-Item $p -Stream *` for `Zone.Identifier` and `Unblock-File $p` only after the user agrees |
| `GetActiveObject` missing | Running in PowerShell 7 | Use `powershell.exe` 5.1, or `[Microsoft.VisualBasic.Interaction]::GetObject($workbookPath)` to target a specific workbook's instance |
| Two XLRisk tabs or functions misbehave | Two registered copies (AddIns folder and repo) | Keep one canonical copy in `%APPDATA%\Microsoft\AddIns` and unregister the other |

## Rules

- Do **not** use `-ExecutionPolicy Bypass`; locally created scripts run under the default
  `RemoteSigned` policy, and the flag is rejected by the permission classifier.
- Keep a single canonical copy of the `.xlam`. Do not register both the repo file and the
  AddIns copy.
- Test cells belong in a scratch workbook closed without saving. Never write test formulas
  into a user's file.
- Do not close or restart the user's Excel to make the add-in appear; report that a restart
  is needed and let them decide.

## Manual fallback (if the script cannot run)

```powershell
powershell.exe -NoProfile -Command '& { $p = Join-Path $env:APPDATA "Microsoft\AddIns\XLRisk.xlam"; $k = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"; Set-ItemProperty $k -Name OPEN -Value ("""" + $p + """"); $app = New-Object -ComObject Excel.Application; $ai = $null; foreach ($a in $app.AddIns) { if ($a.Name -like "XLRisk*") { $ai = $a } }; if ($ai.Installed) { $ai.Installed = $false }; $ai.Installed = $true; $wb = $app.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Cells.Item(1,1).Formula = "=RiskUniform(0,1)"; $app.Calculate(); $ws.Cells.Item(1,1).Text; $wb.Close($false); $app.Quit() }'
```
