[CmdletBinding()]
param(
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$rootDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutputPath) {
    $OutputPath = Join-Path $rootDir "dist\Install Uncle Bao Council Single File.cmd"
}

$payloadFiles = @(
    "UncleBaoCouncil.ps1",
    "Launch Uncle Bao Council.cmd",
    "Uninstall Uncle Bao Council.cmd",
    "council-assets.zip"
)

foreach ($file in $payloadFiles) {
    $fullPath = Join-Path $rootDir $file
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Missing payload file: $file"
    }
}

$tempRoot = Join-Path $env:TEMP ("uncle-bao-council-build-" + [Guid]::NewGuid().ToString("N"))
$stageDir = Join-Path $tempRoot "stage"
$zipPath = Join-Path $tempRoot "payload.zip"

try {
    New-Item -Path $stageDir -ItemType Directory -Force | Out-Null

    foreach ($file in $payloadFiles) {
        Copy-Item -LiteralPath (Join-Path $rootDir $file) -Destination (Join-Path $stageDir $file) -Force
    }

    Compress-Archive -LiteralPath ($payloadFiles | ForEach-Object { Join-Path $stageDir $_ }) -DestinationPath $zipPath -Force
    $payloadBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($zipPath))
    $payloadLines = [regex]::Matches($payloadBase64, ".{1,120}") | ForEach-Object { $_.Value }

    $extractCommand = '$selfPath = $env:SELF; $marker = "__UNCLE_BAO_PAYLOAD__"; $raw = [System.IO.File]::ReadAllText($selfPath); $markerIndex = $raw.IndexOf($marker); if ($markerIndex -lt 0) { throw "Payload marker missing." }; $payload = $raw.Substring($markerIndex + $marker.Length).Trim() -replace "\s", ""; $bytes = [Convert]::FromBase64String($payload); $tempRoot = Join-Path $env:TEMP ("UncleBaoCouncilBundle-" + [Guid]::NewGuid().ToString("N")); New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null; try { $payloadZipPath = Join-Path $tempRoot "payload.zip"; [System.IO.File]::WriteAllBytes($payloadZipPath, $bytes); Expand-Archive -LiteralPath $payloadZipPath -DestinationPath $tempRoot -Force; & (Join-Path $tempRoot "UncleBaoCouncil.ps1") -Action Install; exit $LASTEXITCODE } finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }'

    $bundleLines = @(
        "@echo off",
        "setlocal",
        "set ""SELF=%~f0""",
        "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""$extractCommand""",
        "exit /b %ERRORLEVEL%",
        "__UNCLE_BAO_PAYLOAD__"
    ) + $payloadLines

    $outputDir = Split-Path -Parent $OutputPath
    if ($outputDir) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    Set-Content -LiteralPath $OutputPath -Value ($bundleLines -join "`r`n") -Encoding ASCII
    Write-Host "Built single-file installer: $OutputPath"
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
