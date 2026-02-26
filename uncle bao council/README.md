# Uncle Bao Council

This app runs in the Windows tray and rolls a random check every second:

- 1 in 1000 chance each second to trigger a jumpscare.
- When triggered, it pops up the picture, plays the sound, fades out, then resumes rolling.
- Tray window text: `the council is deciding`.

## Files

- `UncleBaoCouncil.ps1`: main tray/background app
- `launch.vbs`: hidden launcher used at startup
- `install.ps1`: installs to `%LOCALAPPDATA%\OHYEAH\UncleBaoCouncil` and adds startup shortcut
- `uninstall.ps1`: removes startup entry, stops running instance, and removes installed files
- `bao1_480x480.webp`: jumpscare image
- `Flashbang Sound Effect (HD)  How to.mp3`: jumpscare sound

## Install

From this folder in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

## Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```
