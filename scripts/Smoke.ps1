#requires -Version 7.4
<#
.SYNOPSIS
  Runs the demo smoke tests (main window, tray agent, fatal error handling) against a published SharePointDiff.exe.
.DESCRIPTION
  Every run uses isolated settings under tmp\smoke-<guid>\ (SPDIFF_SETTINGS_PATH), synthetic data and no network.
  The script checks exit codes and check files, and that nothing outside tmp\ changed: the default settings file,
  the HKCU Run value and *_SPDiff data folders in the watched folder.
#>
[CmdletBinding()]
param(
    [string]$Exe = (Join-Path $PSScriptRoot '..\publish\win-x64\SharePointDiff.exe'),
    [string]$WatchFolder = (Split-Path (Split-Path (Resolve-Path (Join-Path $PSScriptRoot '..')) -Parent) -Parent),
    [int]$TimeoutSeconds = 240
)
$ErrorActionPreference = 'Stop'
$repo = Resolve-Path (Join-Path $PSScriptRoot '..')
$exePath = (Resolve-Path $Exe).Path
$work = Join-Path $repo ("tmp\smoke-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null
$settings = Join-Path $work 'settings\settings.json'
$logs = Join-Path $work 'settings\logs'

function Get-Tree([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return 'absent' }
    ((Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | ForEach-Object { "$($_.FullName)|$($_.Length)|$($_.LastWriteTimeUtc.Ticks)" }) | Sort-Object) -join "`n"
}
function Get-Snapshot {
    [pscustomobject]@{
        # Real app data, every *_SPDiff data folder next to real mirrors (recursively), and the demo's fake mirror root.
        AppData     = Get-Tree (Join-Path $env:LOCALAPPDATA 'SharePointDiff')
        DataFolders = ((Get-ChildItem -LiteralPath $WatchFolder -Directory -Filter '*_SPDiff' -ErrorAction SilentlyContinue) | ForEach-Object { $_.FullName + "`n" + (Get-Tree $_.FullName) }) -join "`n"
        DemoRoot    = (Test-Path -LiteralPath 'D:\SharePointMirror') -or (Test-Path -LiteralPath 'D:\SharePointMirror_SPDiff')
        RunValue    = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name SharePointDiff -ErrorAction SilentlyContinue).SharePointDiff
    }
}
function Invoke-Run([string]$Name, [string[]]$Arguments, [int]$ExpectedExit, [string]$ChecksFile) {
    $before = @(Get-ChildItem -LiteralPath $logs -Filter 'error-*.txt' -ErrorAction SilentlyContinue).Count
    $process = Start-Process -FilePath $exePath -ArgumentList $Arguments -PassThru -Environment @{ SPDIFF_SETTINGS_PATH = $settings }
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) { $process.Kill($true); return [pscustomobject]@{ Run = $Name; Passed = $false; Detail = "timed out after $TimeoutSeconds s" } }
    $detail = "exit $($process.ExitCode) (expected $ExpectedExit)"; $passed = $process.ExitCode -eq $ExpectedExit
    if ($ChecksFile) {
        if (Test-Path -LiteralPath $ChecksFile) { $checks = Get-Content -LiteralPath $ChecksFile -Raw | ConvertFrom-Json; $passed = $passed -and $checks.Passed; if (-not $checks.Passed) { $detail += "; " + $checks.Error } }
        else { $passed = $false; $detail += '; no check file' }
    }
    $newLogs = @(Get-ChildItem -LiteralPath $logs -Filter 'error-*.txt' -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc | Select-Object -Skip $before)
    if ($ExpectedExit -eq 1) {
        if ($newLogs.Count -eq 0) { $passed = $false; $detail += '; no error log written' }
        elseif (-not (Select-String -LiteralPath $newLogs[-1].FullName -Pattern 'FATAL_SMOKE' -Quiet)) { $passed = $false; $detail += '; the error log is not the deliberate fault' }
    }
    elseif ($newLogs.Count -gt 0) { $passed = $false; $detail += "; unexpected error log $($newLogs[-1].Name)" }
    [pscustomobject]@{ Run = $Name; Passed = $passed; Detail = $detail }
}

$snapshot = Get-Snapshot
$results = @(
    Invoke-Run 'main window'  @('--demo', '--ui-smoke', "`"$work\main`"")            0 "$work\main\ui-checks.json"
    Invoke-Run 'tray agent'   @('--demo', '--agent', '--ui-smoke', "`"$work\agent`"") 0 "$work\agent\agent-checks.json"
    Invoke-Run 'fatal (main)' @('--demo', '--fatal-smoke')                           1 $null
    Invoke-Run 'fatal (tray)' @('--demo', '--agent', '--fatal-smoke', 'tray')        1 $null
)
$unchanged = ($snapshot | ConvertTo-Json) -eq (Get-Snapshot | ConvertTo-Json)
# Demo runs must not even create the redirected settings or a token cache.
$isolated = -not (Test-Path -LiteralPath $settings) -and -not (Test-Path -LiteralPath (Join-Path $work 'settings\msal'))
$results += [pscustomobject]@{ Run = 'no settings or token cache created'; Passed = $isolated; Detail = if ($isolated) { 'demo wrote neither settings.json nor msal\' } else { 'demo created settings or a token cache' } }
$results += [pscustomobject]@{ Run = 'nothing outside tmp changed'; Passed = $unchanged; Detail = if ($unchanged) { 'app data, *_SPDiff folders, demo root and Run value unchanged' } else { 'CHANGED: ' + ((Get-Snapshot) | ConvertTo-Json -Compress) } }
$results | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
Write-Host "Output: $work"
if ($results.Passed -contains $false) { exit 1 }
