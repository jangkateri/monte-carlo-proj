<#
.SYNOPSIS
    Installs and verifies the XLRisk Monte Carlo add-in for desktop Excel (Windows).

.DESCRIPTION
    Places XLRisk.xlam in the per-user Excel add-ins folder, registers it through the
    registry key Excel itself uses for autoload (HKCU ... Excel\Options, OPEN value),
    forces a reload, and checks that the Risk* worksheet functions evaluate.

    Must run under Windows PowerShell 5.1 (powershell.exe). PowerShell 7 cannot attach
    to a running Excel instance.

    Registration uses the registry rather than Application.AddIns.Add, because AddIns.Add
    fails with "Unable to get the Add property of the AddIns class" when Excel is started
    by automation.

.PARAMETER XlamPath
    Path to the XLRisk.xlam build. Defaults to the local repo copy.

.PARAMETER NoVerify
    Skip the formula verification step.

.EXAMPLE
    powershell.exe -NoProfile -File install-xlrisk.ps1
    powershell.exe -NoProfile -File install-xlrisk.ps1 -XlamPath "D:\builds\XLRisk.xlam"
#>
[CmdletBinding()]
param(
    [string]$XlamPath = (Join-Path $env:USERPROFILE 'monte-carlo-sim\XLRisk\XLRisk.xlam'),
    [switch]$NoVerify
)

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'Run with Windows PowerShell 5.1: powershell.exe -NoProfile -File "<this script>"'
}

function Write-Step([string]$Message) { Write-Host "[xlrisk] $Message" }

function Get-XlRiskAddIn($App) {
    foreach ($addIn in $App.AddIns) {
        if ($addIn.Name -like 'XLRisk*') { return $addIn }
    }
    return $null
}

function Get-OpenValues([string]$Version) {
    $regPath = "HKCU:\Software\Microsoft\Office\$Version\Excel\Options"
    if (-not (Test-Path $regPath)) { return @() }
    $props = Get-ItemProperty $regPath
    return @($props.PSObject.Properties |
        Where-Object { $_.Name -match '^OPEN\d*$' } |
        ForEach-Object { [pscustomobject]@{ Name = $_.Name; Value = [string]$_.Value } })
}

function Test-SamePath([string]$A, [string]$B) {
    if (-not $A -or -not $B) { return $false }
    return [string]::Equals($A.Trim('"'), $B.Trim('"'), [StringComparison]::OrdinalIgnoreCase)
}

function Set-XlRiskRegistration([string]$Version, [string]$Path) {
    $regPath = "HKCU:\Software\Microsoft\Office\$Version\Excel\Options"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    $current = Get-OpenValues -Version $Version
    $existing = $current | Where-Object { $_.Value -like '*XLRisk*' } | Select-Object -First 1
    if ($existing) {
        return [pscustomobject]@{ Name = $existing.Name; Value = $existing.Value; Created = $false }
    }

    $slot = 'OPEN'
    $n = 1
    while ($current.Name -contains $slot) {
        $slot = "OPEN$n"
        $n++
    }
    $value = '"' + $Path + '"'
    Set-ItemProperty -Path $regPath -Name $slot -Value $value
    return [pscustomobject]@{ Name = $slot; Value = $value; Created = $true }
}

function Invoke-XlRiskVerification {
    param($App)

    # XLRisk returns the deterministic mean while no simulation runs.
    $tests = @(
        [pscustomobject]@{ Formula = '=RiskUniform(0,1)';   Expected = 0.5 },
        [pscustomobject]@{ Formula = '=RiskNormal(100,15)'; Expected = 100.0 }
    )

    $wb = $App.Workbooks.Add()
    try {
        $ws = $wb.Worksheets.Item(1)
        for ($i = 0; $i -lt $tests.Count; $i++) {
            $ws.Cells.Item($i + 1, 1).Formula = $tests[$i].Formula
        }
        $App.Calculate()

        $results = for ($i = 0; $i -lt $tests.Count; $i++) {
            $cell = $ws.Cells.Item($i + 1, 1)
            $text = [string]$cell.Text
            $pass = $false
            if ($text -notlike '#*') {
                try {
                    $value = [double]$cell.Value2
                    $pass = [math]::Abs($value - $tests[$i].Expected) -lt 0.0001
                } catch { }
            }
            [pscustomobject]@{
                Formula  = $tests[$i].Formula
                Expected = $tests[$i].Expected
                Actual   = $text
                Pass     = $pass
            }
        }
        return $results
    }
    finally {
        $wb.Close($false)
    }
}

$userApp = $null
$userEntry = $null
$exitCode = 1
$verification = @()

try {
    if (-not (Test-Path -LiteralPath $XlamPath)) {
        throw "XLRisk.xlam not found at: $XlamPath"
    }

    # 1. Probe: learn the Excel version and any existing registration.
    $probe = $null
    $createdProbe = $false
    try {
        $probe = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $userApp = $probe
        Write-Step "Attached to running Excel (version $($probe.Version), HWND $($probe.Hwnd))."
    }
    catch {
        $probe = New-Object -ComObject Excel.Application
        $createdProbe = $true
        Write-Step 'Excel was not running; started a hidden instance to inspect.'
    }

    $version = [string]$probe.Version
    if ([double]$version -lt 15) {
        throw "Excel $version is too old; XLRisk requires Excel 2013 (15.0) or later."
    }

    $existing = Get-XlRiskAddIn $probe
    $existingPath = if ($existing) { $existing.FullName } else { $null }
    if ($existing) { Write-Step "Existing add-in entry: $existingPath (Installed=$($existing.Installed))" }

    if ($createdProbe) { $probe.Quit(); $probe = $null }

    # 2. Choose the canonical file location.
    $targetDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
    $standardTarget = Join-Path $targetDir 'XLRisk.xlam'
    $destination = $standardTarget
    if ($existingPath) {
        $parent = Split-Path -Parent $existingPath
        if ($parent -and (Test-Path -LiteralPath $parent)) { $destination = $existingPath }
    }

    # 3. Release the file in the user's running Excel, then copy when needed.
    if ($userApp -and $existing) {
        $userEntry = $existing
        if ($existing.Installed) {
            $existing.Installed = $false
            Start-Sleep -Milliseconds 300
        }
    }

    if (-not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $sourceHash = (Get-FileHash -LiteralPath $XlamPath -Algorithm SHA256).Hash
    $needCopy = $true
    if (Test-Path -LiteralPath $destination) {
        $needCopy = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash -ne $sourceHash
    }
    if ($needCopy) {
        Copy-Item -LiteralPath $XlamPath -Destination $destination -Force
        Write-Step "Copied add-in to $destination"
    }
    else {
        Write-Step "Add-in file already current: $destination"
    }

    # 4. Register for autoload (registry OPEN value, the mechanism Excel itself uses).
    $registration = Set-XlRiskRegistration -Version $version -Path $destination
    if ($registration.Created) {
        Write-Step "Registered autoload: $($registration.Name) = $($registration.Value)"
    }
    else {
        Write-Step "Autoload already registered: $($registration.Name) = $($registration.Value)"
    }

    # 5. Make it live in the user's already-open Excel, when possible.
    if ($userApp) {
        if ($userEntry) {
            if ($userEntry.Installed) { $userEntry.Installed = $false }
            $userEntry.Installed = $true
            Write-Step 'Reloaded the add-in in the running Excel instance.'
        }
        else {
            Write-Warning 'No XLRisk entry in the running Excel. Restart Excel for it to appear there; this script will not close your workbooks.'
        }
    }

    # 6. Verify in a fresh instance started after registration.
    if (-not $NoVerify) {
        Write-Step 'Verifying Risk* functions in a fresh Excel instance...'
        $verifyApp = New-Object -ComObject Excel.Application
        try {
            $entry = $null
            for ($i = 0; $i -lt 20 -and -not $entry; $i++) {
                $entry = Get-XlRiskAddIn $verifyApp
                if (-not $entry) { Start-Sleep -Milliseconds 250 }
            }
            if (-not $entry) {
                throw "XLRisk did not appear in the add-in list. Registration may be blocked by Trust Center or a policy."
            }
            # Excel can load an add-in at startup without initialising its VBA; a reload fixes it.
            if ($entry.Installed) { $entry.Installed = $false; Start-Sleep -Milliseconds 300 }
            $entry.Installed = $true
            Start-Sleep -Milliseconds 300

            $verification = @(Invoke-XlRiskVerification -App $verifyApp)
        }
        finally {
            try { $verifyApp.Quit() } catch { }
        }
        $verification | Format-Table -AutoSize | Out-String | Write-Host
    }

    $failed = @($verification | Where-Object { -not $_.Pass })
    if ($verification.Count -gt 0 -and $failed.Count -gt 0) {
        Write-Step 'Verification FAILED - the add-in is registered but not evaluating.'
        $exitCode = 1
    }
    elseif ($verification.Count -gt 0) {
        Write-Step 'Verification passed - XLRisk is installed and working.'
        $exitCode = 0
    }
    else {
        Write-Step 'Install complete (verification skipped).'
        $exitCode = 0
    }
}
finally {
    if ($probe -and $createdProbe) {
        try { $probe.Quit() } catch { }
    }
}

exit $exitCode
