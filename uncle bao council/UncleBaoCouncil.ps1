Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    throw "This script requires Windows Forms and System.Drawing."
}

try {
    Add-Type -AssemblyName PresentationCore -ErrorAction Stop
} catch {
    # PresentationCore may be unavailable; sound playback will fall back.
}

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$script:AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SourceImagePath = Join-Path $script:AppDir "bao1_480x480.webp"
$script:ConvertedImagePath = Join-Path $script:AppDir "bao1_480x480.converted.png"
$script:ImagePath = $script:SourceImagePath
$script:SoundPath = Join-Path $script:AppDir "Flashbang Sound Effect (HD)  How to.mp3"

$script:Random = [System.Random]::new()
$script:ChanceDenominator = 1000
$script:InitialSoundVolume = 100
$script:FadeDurationMs = 4000
$script:FadeIntervalMs = 50
$script:SoundLeadInMs = 600
$script:ScareActive = $false
$script:IsExiting = $false
$script:AppContext = $null
$script:PollTimer = $null
$script:StatusForm = $null
$script:NotifyIcon = $null
$script:PreloadedSoundPlayer = $null

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

function Ensure-AssetFilesHidden {
    Set-HiddenSystemFileAttribute -Path $script:SourceImagePath
    Set-HiddenSystemFileAttribute -Path $script:ConvertedImagePath
    Set-HiddenSystemFileAttribute -Path $script:SoundPath
}

function Start-ScareSound {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    if ($null -ne $script:PreloadedSoundPlayer) {
        try {
            $script:PreloadedSoundPlayer.settings.volume = $script:InitialSoundVolume
            $script:PreloadedSoundPlayer.controls.stop() | Out-Null
            $script:PreloadedSoundPlayer.controls.currentPosition = 0
            $script:PreloadedSoundPlayer.controls.play() | Out-Null

            return [PSCustomObject]@{
                Kind = "WmPlayerShared"
                Ref  = $script:PreloadedSoundPlayer
            }
        } catch {
            # Fall through to non-shared playback methods.
        }
    }

    try {
        $mediaPlayer = New-Object System.Windows.Media.MediaPlayer
        $mediaPlayer.Open([System.Uri]::new($Path))
        $mediaPlayer.Volume = 1.0
        $mediaPlayer.Play()

        return [PSCustomObject]@{
            Kind = "MediaPlayer"
            Ref  = $mediaPlayer
        }
    } catch {
        # Fallback to Windows Media Player COM object when WPF media is unavailable.
    }

    try {
        $wmPlayer = New-Object -ComObject WMPlayer.OCX
        $wmPlayer.settings.volume = 100
        $wmPlayer.URL = $Path
        $wmPlayer.controls.play()

        return [PSCustomObject]@{
            Kind = "WmPlayer"
            Ref  = $wmPlayer
        }
    } catch {
        return $null
    }
}

function Stop-ScareSound {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Handle
    )

    if ($null -eq $Handle) {
        return
    }

    if ($Handle.Kind -eq "MediaPlayer" -and $null -ne $Handle.Ref) {
        try {
            $Handle.Ref.Stop()
            $Handle.Ref.Close()
        } catch {
            # Ignore cleanup issues.
        }
        return
    }

    if ($Handle.Kind -eq "WmPlayerShared" -and $null -ne $Handle.Ref) {
        try {
            $Handle.Ref.controls.stop() | Out-Null
            $Handle.Ref.settings.volume = $script:InitialSoundVolume
        } catch {
            # Ignore cleanup issues.
        }
        return
    }

    if ($Handle.Kind -eq "WmPlayer" -and $null -ne $Handle.Ref) {
        try {
            $Handle.Ref.controls.stop() | Out-Null
        } catch {
            # Ignore cleanup issues.
        }

        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Handle.Ref)
        } catch {
            # Ignore cleanup issues.
        }
    }
}

function Set-ScareSoundVolume {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Handle,

        [Parameter(Mandatory = $true)]
        [int]$VolumePercent
    )

    if ($null -eq $Handle) {
        return
    }

    $clampedVolume = [Math]::Max(0, [Math]::Min(100, $VolumePercent))

    if ($Handle.Kind -eq "MediaPlayer" -and $null -ne $Handle.Ref) {
        try {
            $Handle.Ref.Volume = ($clampedVolume / 100.0)
        } catch {
            # Ignore volume update errors.
        }
        return
    }

    if (($Handle.Kind -eq "WmPlayerShared" -or $Handle.Kind -eq "WmPlayer") -and $null -ne $Handle.Ref) {
        try {
            $Handle.Ref.settings.volume = $clampedVolume
        } catch {
            # Ignore volume update errors.
        }
    }
}

function Initialize-ScareSound {
    if (-not (Test-Path -LiteralPath $script:SoundPath)) {
        return
    }

    try {
        $wmPlayer = New-Object -ComObject WMPlayer.OCX
        $wmPlayer.settings.autoStart = $false
        $wmPlayer.settings.volume = $script:InitialSoundVolume
        $wmPlayer.URL = $script:SoundPath

        # Wait briefly for media open and prime decoding to reduce first-play latency.
        for ($i = 0; $i -lt 40; $i++) {
            try {
                if ([int]$wmPlayer.openState -ge 13) {
                    break
                }
            } catch {
                break
            }
            Start-Sleep -Milliseconds 25
        }

        try {
            $wmPlayer.settings.volume = 0
            $wmPlayer.controls.play() | Out-Null
            Start-Sleep -Milliseconds 140
            $wmPlayer.controls.pause() | Out-Null
            $wmPlayer.controls.currentPosition = 0
        } catch {
            # Ignore priming errors.
        } finally {
            $wmPlayer.settings.volume = $script:InitialSoundVolume
            $wmPlayer.controls.stop() | Out-Null
        }

        $script:PreloadedSoundPlayer = $wmPlayer
    } catch {
        $script:PreloadedSoundPlayer = $null
    }
}

function Get-ScareImage {
    if (-not (Test-Path -LiteralPath $script:SourceImagePath)) {
        return $null
    }

    if (Test-Path -LiteralPath $script:ConvertedImagePath) {
        try {
            $cachedImage = [System.Drawing.Image]::FromFile($script:ConvertedImagePath)
            try {
                return New-Object System.Drawing.Bitmap $cachedImage
            } finally {
                $cachedImage.Dispose()
            }
        } catch {
            # Keep trying with source file.
        }
    }

    $extension = [System.IO.Path]::GetExtension($script:SourceImagePath).ToLowerInvariant()
    if ($extension -eq ".webp") {
        try {
            $uri = [System.Uri]::new($script:SourceImagePath)
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
                $pngBytes = $memoryStream.ToArray()
            } finally {
                $memoryStream.Dispose()
            }

            try {
                [System.IO.File]::WriteAllBytes($script:ConvertedImagePath, $pngBytes)
            } catch {
                # Ignore cache write failures.
            }
            Set-HiddenSystemFileAttribute -Path $script:ConvertedImagePath

            $imageStream = [System.IO.MemoryStream]::new($pngBytes)
            try {
                $decodedImage = [System.Drawing.Image]::FromStream($imageStream)
                try {
                    return New-Object System.Drawing.Bitmap $decodedImage
                } finally {
                    $decodedImage.Dispose()
                }
            } finally {
                $imageStream.Dispose()
            }
        } catch {
            return $null
        }
    }

    try {
        $sourceImage = [System.Drawing.Image]::FromFile($script:SourceImagePath)
        try {
            return New-Object System.Drawing.Bitmap $sourceImage
        } finally {
            $sourceImage.Dispose()
        }
    } catch {
        return $null
    }
}

function Show-StatusWindow {
    if ($null -eq $script:StatusForm -or $script:IsExiting) {
        return
    }

    $script:StatusForm.Show()
    $script:StatusForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $script:StatusForm.BringToFront()
    $script:StatusForm.Activate()
}

function Show-JumpScare {
    if ($script:ScareActive -or $script:IsExiting) {
        return
    }

    $script:ScareActive = $true

    try {
        $loadedImage = Get-ScareImage

        if ($null -eq $loadedImage) {
            $script:ScareActive = $false
            return
        }

        [void](Start-ScareSound -Path $script:SoundPath)

        $scareForm = New-Object System.Windows.Forms.Form
        $scareForm.Text = "UNCLE BAO COUNCIL"
        $scareForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
        $scareForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
        $scareForm.Bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $scareForm.BackColor = [System.Drawing.Color]::Black
        $scareForm.TopMost = $true
        $scareForm.Opacity = 1.0
        $scareForm.ShowInTaskbar = $false

        $container = New-Object System.Windows.Forms.Panel
        $container.Dock = [System.Windows.Forms.DockStyle]::Fill
        $container.BackColor = [System.Drawing.Color]::Black

        $picture = New-Object System.Windows.Forms.PictureBox
        $picture.Dock = [System.Windows.Forms.DockStyle]::Fill
        $picture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $picture.BackColor = [System.Drawing.Color]::Black
        $picture.Image = $loadedImage
        $container.Controls.Add($picture)
        $scareForm.Controls.Add($container)

        $leadInTimer = New-Object System.Windows.Forms.Timer
        $leadInTimer.Interval = [Math]::Max(1, [int]$script:SoundLeadInMs)

        $fadeTimer = New-Object System.Windows.Forms.Timer
        $fadeTimer.Interval = $script:FadeIntervalMs
        $fadeState = [PSCustomObject]@{
            Steps = [Math]::Max(1, [int]([Math]::Round($script:FadeDurationMs / [double]$script:FadeIntervalMs)))
            Step  = 0
        }

        $leadInTimer.Add_Tick({
            try {
                $leadInTimer.Stop()

                if ($null -eq $scareForm -or $scareForm.IsDisposed) {
                    return
                }

                if (-not $scareForm.Visible) {
                    $scareForm.Show()
                }

                $fadeTimer.Start()
            } catch {
                try { $leadInTimer.Stop() } catch {}
                try { if ($null -ne $scareForm -and -not $scareForm.IsDisposed) { $scareForm.Close() } } catch {}
            }
        }.GetNewClosure())

        $fadeTimer.Add_Tick({
            try {
                if ($null -eq $scareForm -or $scareForm.IsDisposed) {
                    $fadeTimer.Stop()
                    return
                }

                $fadeState.Step = [int]$fadeState.Step + 1
                $progress = [Math]::Min(1.0, ($fadeState.Step / [double]$fadeState.Steps))
                $nextOpacity = [Math]::Max(0.0, (1.0 - $progress))
                $scareForm.Opacity = $nextOpacity

                if ($progress -ge 1.0) {
                    $fadeTimer.Stop()
                    $scareForm.Close()
                    return
                }
            } catch {
                try { $fadeTimer.Stop() } catch {}
                try { if ($null -ne $scareForm -and -not $scareForm.IsDisposed) { $scareForm.Close() } } catch {}
            }
        }.GetNewClosure())

        $scareForm.Add_FormClosed({
            try { $leadInTimer.Stop() } catch {}
            try { $fadeTimer.Stop() } catch {}
            try { $leadInTimer.Dispose() } catch {}
            try { $fadeTimer.Dispose() } catch {}

            if ($picture.Image -is [System.Drawing.Image]) {
                try { $picture.Image.Dispose() } catch {}
                $picture.Image = $null
            }

            $script:ScareActive = $false
        }.GetNewClosure())

        if ($script:SoundLeadInMs -le 0) {
            $scareForm.Show()
            $fadeTimer.Start()
        } else {
            $leadInTimer.Start()
        }
    } catch {
        $script:ScareActive = $false
    }
}

function Stop-App {
    if ($script:IsExiting) {
        return
    }

    $script:IsExiting = $true

    if ($null -ne $script:PollTimer) {
        try {
            $script:PollTimer.Stop()
            $script:PollTimer.Dispose()
        } catch {
            # Ignore cleanup issues.
        }
    }

    if ($null -ne $script:NotifyIcon) {
        try {
            $script:NotifyIcon.Visible = $false
            $script:NotifyIcon.Dispose()
        } catch {
            # Ignore cleanup issues.
        }
    }

    if ($null -ne $script:StatusForm) {
        try {
            $script:StatusForm.Close()
            $script:StatusForm.Dispose()
        } catch {
            # Ignore cleanup issues.
        }
    }

    if ($null -ne $script:PreloadedSoundPlayer) {
        try {
            $script:PreloadedSoundPlayer.controls.stop() | Out-Null
        } catch {
            # Ignore cleanup issues.
        }

        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:PreloadedSoundPlayer)
        } catch {
            # Ignore cleanup issues.
        }

        $script:PreloadedSoundPlayer = $null
    }

    if ($null -ne $script:AppContext) {
        try {
            $script:AppContext.ExitThread()
        } catch {
            # Ignore cleanup issues.
        }
    }
}

$script:StatusForm = New-Object System.Windows.Forms.Form
$script:StatusForm.Text = "Uncle Bao Council"
$script:StatusForm.Size = New-Object System.Drawing.Size(360, 140)
$script:StatusForm.MinimumSize = New-Object System.Drawing.Size(360, 140)
$script:StatusForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$script:StatusForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$script:StatusForm.MaximizeBox = $false
$script:StatusForm.MinimizeBox = $true

$messageLabel = New-Object System.Windows.Forms.Label
$messageLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$messageLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$messageLabel.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 16, [System.Drawing.FontStyle]::Regular)
$messageLabel.Text = "the council is deciding"
$script:StatusForm.Controls.Add($messageLabel)

$script:StatusForm.Add_FormClosing({
    param($sender, $eventArgs)

    if (-not $script:IsExiting -and $eventArgs.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
        $eventArgs.Cancel = $true
        $sender.Hide()
    }
})

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$openMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openMenuItem.Text = "Open Council Window"
$openMenuItem.Add_Click({ Show-StatusWindow })
[void]$trayMenu.Items.Add($openMenuItem)

$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.Add_Click({ Stop-App })
[void]$trayMenu.Items.Add($exitMenuItem)

$script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
$script:NotifyIcon.Icon = [System.Drawing.SystemIcons]::Information
$script:NotifyIcon.Visible = $true
$script:NotifyIcon.Text = "Uncle Bao Council"
$script:NotifyIcon.ContextMenuStrip = $trayMenu
$script:NotifyIcon.Add_DoubleClick({ Show-StatusWindow })

Ensure-AssetFilesHidden
Initialize-ScareSound

$script:PollTimer = New-Object System.Windows.Forms.Timer
$script:PollTimer.Interval = 1000
$script:PollTimer.Add_Tick({
    if ($script:IsExiting -or $script:ScareActive) {
        return
    }

    try {
        if ($script:Random.Next(1, ($script:ChanceDenominator + 1)) -eq 1) {
            Show-JumpScare
        }
    } catch {
        # Keep polling even if a single jumpscare attempt fails.
    }
})
$script:PollTimer.Start()

$script:NotifyIcon.ShowBalloonTip(
    2500,
    "Uncle Bao Council",
    "the council is deciding",
    [System.Windows.Forms.ToolTipIcon]::Info
)

$script:AppContext = New-Object System.Windows.Forms.ApplicationContext

try {
    [System.Windows.Forms.Application]::Run($script:AppContext)
} finally {
    Stop-App
}
