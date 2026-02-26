[CmdletBinding()]
param(
    [switch]$KeepFiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$installDir = Join-Path $env:LOCALAPPDATA "OHYEAH\UncleBaoCouncil"
$startupDir = [Environment]::GetFolderPath("Startup")
$startupShortcutPath = Join-Path $startupDir "Uncle Bao Council.lnk"

$runningProcesses = Get-CimInstance Win32_Process | Where-Object {
    ($_.Name -ieq "powershell.exe" -or $_.Name -ieq "pwsh.exe" -or $_.Name -ieq "wscript.exe") -and
    $_.CommandLine -and
    $_.CommandLine -like "*UncleBaoCouncil*"
}

foreach ($process in $runningProcesses) {
    try {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction Stop
    } catch {
        # Ignore already-exited processes.
    }
}

if (Test-Path -LiteralPath $startupShortcutPath) {
    Remove-Item -LiteralPath $startupShortcutPath -Force
}

if (-not $KeepFiles -and (Test-Path -LiteralPath $installDir)) {
    Remove-Item -LiteralPath $installDir -Recurse -Force
}

Write-Host ""
Write-Host "Uncle Bao Council removed."
if ($KeepFiles) {
    Write-Host "Install folder kept at: $installDir"
}
