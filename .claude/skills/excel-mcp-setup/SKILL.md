---
name: excel-mcp-setup
description: Install and register the Excel MCP server (@sbroenne/mcp-server-excel) in Claude Code on Windows when it has never been installed, and verify it connects. Use when the user asks to install or set up the Excel MCP server, when mcp__excel-mcp__* tools are missing, or when the server shows Failed to connect.
---

# Excel MCP Server Setup (Windows)

Registers `@sbroenne/mcp-server-excel` as a stdio MCP server in Claude Code and proves it
connects. This is an idempotent install: if the server is already configured, report its
status and change nothing.

## Hard requirements

- Windows 10/11 with an **interactive desktop** session (not a headless service).
- **Desktop Microsoft Excel 2016 or later** installed. Excel for the web will not work.
- **Node.js 18+** with `npx` on PATH.
- Network access on first run - `npx` downloads the package and its native runtime.
- No .NET SDK or runtime is needed; the npm package ships a self-contained binary through
  platform optional dependencies (`@sbroenne/mcp-server-excel-win32-x64` / `-win32-arm64`).
  Do not disable optional dependencies.

## Decision flow

1. **Check whether it is already installed:**
   ```
   claude mcp get excel-mcp
   ```
   - Prints `excel-mcp:`, `Scope:`, `Status: ✔ Connected`, `Command: npx` -> already installed.
     If the status is Connected, **stop**; do not re-add it.
   - `No MCP server named "excel-mcp". Configured servers: ...` -> never installed, continue.
   - `claude mcp list` shows every server and its health.
2. **Run the install script** (safe to run repeatedly; it skips an existing registration):
   ```
   powershell.exe -NoProfile -File "C:\Users\John Ang\monte-carlo-proj\.claude\skills\excel-mcp-setup\scripts\install-excel-mcp.ps1"
   ```
   Optional: `-Scope project` writes `.mcp.json` for a team; `-Scope local` keeps it to the
   current project. The default `user` scope matches this machine's existing setup.
3. **Equivalent manual command** if the script cannot run:
   ```
   claude mcp add --scope user excel-mcp -- npx -y @sbroenne/mcp-server-excel@latest
   ```
4. **Verify** with `claude mcp get excel-mcp` (Status must say Connected) and tell the user
   to **restart Claude Code**: MCP servers load at session start, so `mcp__excel-mcp__*`
   tools do not appear in the current session.

## What a correct registration looks like

In `~/.claude.json` (user scope) under the top-level `mcpServers` key:

```json
"excel-mcp": {
  "command": "npx",
  "args": ["-y", "@sbroenne/mcp-server-excel@latest"],
  "env": {}
}
```

`claude mcp get excel-mcp` then reports `Scope: User config (available in all your projects)`
and `Status: ✔ Connected`.

Removal: `claude mcp remove excel-mcp -s user`.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `No MCP server named "excel-mcp"` right after adding | Added under a different scope/name | `claude mcp list`; inspect `~/.claude.json` or `.mcp.json`; re-add with the intended `--scope` |
| `Failed to connect - CONNECTION_CLOSED` | Server process crashed or cannot start | Run `npx -y @sbroenne/mcp-server-excel@latest` in a terminal and read stderr |
| Connect fails, `EAI_AGAIN` / `ETIMEDOUT` / certificate error | Proxy or registry blocking npx download | Check `npm config get proxy`, `npm config get https-proxy`, `npm config get registry`; set `NODE_EXTRA_CA_CERTS` for corporate TLS |
| `No matching native package` / architecture error | Node arch has no matching optional dependency (or optional deps disabled) | Check `node -p process.arch`; reinstall Node matching the platform; keep optional deps enabled |
| `node: command not found`, or version < 18 | Missing/old Node.js | Install Node.js 18+ (LTS recommended) and reopen the terminal |
| Server connected but no `mcp__excel-mcp__*` tools | Claude Code has not restarted since registration | Restart Claude Code, start a new session |
| Tools fail with COM / `Class not registered` | Desktop Excel missing or Office damaged | Verify Excel opens manually; this server does not support Excel for the web |
| Excel shows a modal dialog and calls time out | Excel is waiting on user input | Respond in the Excel window; the server reports `dialogOpen` and `canClose:false` |
| Workbook opens read-only or "file in use" | The file is already open in another Excel window | Use `file_read discover` + `file attach` instead of reopening |
| `Risk*` formulas or add-in macros do nothing (`#NAME?`) | MCP sessions start Excel with `AutomationSecurity=3`, which disables all macros, including add-ins | Model in a normally opened Excel window; for XLRisk use the `xlrisk-setup` skill |
| Server runs but Excel never appears | Expected: the server may drive Excel hidden | Pass `show:true` on open/create only when the user wants to see Excel |
| Works on a workstation but not over RDP/service | No interactive desktop session | Run in a signed-in desktop session, not a headless service |

## Rules

- Check `claude mcp get excel-mcp` first; never register a duplicate name.
- Do not change an existing user-scope entry's command or version without asking; report it.
- Do not run `npm install -g` unless `npx` is unavailable; the registered command is `npx`.
- `-y` on `npx` is required for non-interactive first-run download.
- Do not restart or kill the user's Claude Code session; tell them a restart is needed.
- Registration writes Claude config (`~/.claude.json` or `.mcp.json`); treat it as a
  user-visible configuration change and report exactly what changed.
