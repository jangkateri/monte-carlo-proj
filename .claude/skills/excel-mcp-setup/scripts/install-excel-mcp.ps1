<#
.SYNOPSIS
    Installs and verifies the Excel MCP server (@sbroenne/mcp-server-excel) in Claude Code.

.DESCRIPTION
    Idempotent: checks for an existing registration first and changes nothing when the
    server is already configured and connected. Otherwise registers it with the Claude
    Code CLI and verifies the connection.

    Requires Windows, desktop Excel 2016+, and Node.js 18+ with npx. The npm package ships
    a self-contained native binary; no .NET runtime is needed.

.PARAMETER Scope
    Claude Code configuration scope: user (default), project, or local.

.PARAMETER Name
    MCP server name. Defaults to excel-mcp.

.EXAMPLE
    powershell.exe -NoProfile -File install-excel-mcp.ps1
    powershell.exe -NoProfile -File install-excel-mcp.ps1 -Scope project
#>
[CmdletBinding()]
param(
    [ValidateSet('user', 'project', 'local')]
    [string]$Scope = 'user',
    [string]$Name = 'excel-mcp'
)

$ErrorActionPreference = 'Stop'

function Write-Step([string]$Message) { Write-Host "[excel-mcp] $Message" }

# --- Prerequisites -----------------------------------------------------------
if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
    throw 'Windows is required: this server drives desktop Excel through COM.'
}

$node = Get-Command node -ErrorAction SilentlyContinue
$npx = Get-Command npx -ErrorAction SilentlyContinue
if (-not $node) { throw 'Node.js 18+ is required, but node was not found on PATH.' }
if (-not $npx) { throw 'npx was not found on PATH. Install Node.js 18+ and reopen the terminal.' }

$nodeVersion = (& node --version).Trim().TrimStart('v')
$major = 0
if (-not [int]::TryParse(($nodeVersion -split '\.')[0], [ref]$major) -or $major -lt 18) {
    throw "Node.js $nodeVersion is too old; version 18 or later is required."
}
Write-Step "Node.js $nodeVersion, npx OK."

$excelFound = $false
foreach ($key in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe')) {
    if (Test-Path $key) { $excelFound = $true; break }
}
if (-not $excelFound) {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\EXCEL.EXE'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\EXCEL.EXE')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    $excelFound = @($candidates).Count -gt 0
}
if (-not $excelFound) {
    throw 'Desktop Microsoft Excel 2016+ was not found. Excel for the web is not supported.'
}
Write-Step 'Desktop Excel found.'

$claude = Get-Command claude -ErrorAction SilentlyContinue
if (-not $claude) { throw 'The Claude Code CLI (claude) was not found on PATH.' }

# --- Existing registration ---------------------------------------------------
Write-Step "Checking for an existing '$Name' registration..."
$existing = & claude mcp get $Name 2>&1 | Out-String
$isInstalled = $existing -notmatch 'No MCP server named'

if ($isInstalled) {
    Write-Host $existing.TrimEnd()
    if ($existing -match 'Connected') {
        Write-Step 'Already installed and connected. Nothing to change.'
        exit 0
    }
    Write-Warning 'Registered but not connected. Leaving the entry untouched; see SKILL.md troubleshooting.'
    exit 1
}

# --- Install -----------------------------------------------------------------
Write-Step "Not installed. Registering at '$Scope' scope..."
$addArgs = @('mcp', 'add', '--scope', $Scope, $Name, '--', 'npx', '-y', '@sbroenne/mcp-server-excel@latest')
& claude @addArgs 2>&1 | ForEach-Object { Write-Host $_ }
if ($LASTEXITCODE -ne 0) {
    throw "claude mcp add failed with exit code $LASTEXITCODE."
}

# --- Verify ------------------------------------------------------------------
Write-Step 'Verifying connection (the first run may download the package and take a minute)...'
$result = & claude mcp get $Name 2>&1 | Out-String
Write-Host $result.TrimEnd()

if ($result -match 'Connected') {
    Write-Step 'Installed and connected. Restart Claude Code to load the mcp__excel-mcp__* tools.'
    exit 0
}

Write-Warning 'Registered, but the health check did not report Connected. See SKILL.md troubleshooting.'
exit 1
