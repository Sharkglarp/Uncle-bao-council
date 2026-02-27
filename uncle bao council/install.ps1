[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$targetDir = Join-Path $env:LOCALAPPDATA "OHYEAH\UncleBaoCouncil"
$startupDir = [Environment]::GetFolderPath("Startup")
$startupShortcutPath = Join-Path $startupDir "Uncle Bao Council.lnk"
$wscriptPath = Join-Path $env:WINDIR "System32\wscript.exe"

function Set-HiddenSystemFileAttribute {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    try {
        $item = Get-Item -LiteralPath $Path -Force
        $item.Attributes = ($item.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System)
    } catch {
        # Ignore attribute failures.
    }
}

$requiredFiles = @(
    "UncleBaoCouncil.ps1",
    "launch.vbs",
    "council-assets.zip"
)

foreach ($file in $requiredFiles) {
    $fullPath = Join-Path $sourceDir $file
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Missing required file: $file"
    }
}

New-Item -Path $targetDir -ItemType Directory -Force | Out-Null

foreach ($file in $requiredFiles) {
    Copy-Item -LiteralPath (Join-Path $sourceDir $file) -Destination (Join-Path $targetDir $file) -Force
}

$assetArchivePath = Join-Path $targetDir "council-assets.zip"
if (-not (Test-Path -LiteralPath $assetArchivePath)) {
    throw "Missing asset archive: council-assets.zip"
}

Expand-Archive -LiteralPath $assetArchivePath -DestinationPath $targetDir -Force
Remove-Item -LiteralPath $assetArchivePath -Force

$sourceImagePath = Join-Path $targetDir "bao1_480x480.webp"
$convertedImagePath = Join-Path $targetDir "bao1_480x480.converted.png"
if (Test-Path -LiteralPath $sourceImagePath) {
    try {
        Add-Type -AssemblyName PresentationCore -ErrorAction Stop
        $uri = [System.Uri]::new($sourceImagePath)
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.UriSource = $uri
        $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.EndInit()
        $bitmap.Freeze()

        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        [void]$encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $memoryStream = New-Object System.IO.MemoryStream
        try {
            $encoder.Save($memoryStream)
            [System.IO.File]::WriteAllBytes($convertedImagePath, $memoryStream.ToArray())
        } finally {
            $memoryStream.Dispose()
        }
    } catch {
        # If conversion fails, runtime will try WebP decode again.
    }
}

Set-HiddenSystemFileAttribute -Path (Join-Path $targetDir "bao1_480x480.webp")
Set-HiddenSystemFileAttribute -Path (Join-Path $targetDir "Flashbang Sound Effect (HD)  How to.mp3")
Set-HiddenSystemFileAttribute -Path $convertedImagePath

# Also hide the source-folder assets after install so GitHub-downloaded folders
# are hidden on the user's machine once setup is run.
Set-HiddenSystemFileAttribute -Path (Join-Path $sourceDir "bao1_480x480.webp")
Set-HiddenSystemFileAttribute -Path (Join-Path $sourceDir "Flashbang Sound Effect (HD)  How to.mp3")
Set-HiddenSystemFileAttribute -Path (Join-Path $sourceDir "bao1_480x480.converted.png")

$launcherPath = Join-Path $targetDir "launch.vbs"
$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($startupShortcutPath)
$shortcut.TargetPath = $wscriptPath
$shortcut.Arguments = "`"$launcherPath`""
$shortcut.WorkingDirectory = $targetDir
$shortcut.IconLocation = "$env:WINDIR\System32\shell32.dll,220"
$shortcut.Description = "Starts Uncle Bao Council in the background."
$shortcut.Save()

Start-Process -FilePath $wscriptPath -ArgumentList "`"$launcherPath`""

#fixed appdata link 26/02/27

Write-Host ""
Write-Host "Uncle Bao Council installed."
Write-Host "Install folder: $targetDir"
Write-Host "Startup shortcut: $startupShortcutPath"
Write-Host "Use uninstall.ps1 to remove it."
