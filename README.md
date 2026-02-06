# Civil 3D 2023 Transfer

This repo provides export, import, verification, and post-install helper scripts for Civil 3D 2023 configuration transfer.

## Quick Links (raw)
- Export: https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Export-Civil3D-2023.ps1
- Import: https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Import-Civil3D-2023.ps1

## Recommended Flow
1) **Install Civil 3D 2023** on the target machine and launch once, then close it.
2) **Export** on the source machine (the one with the desired settings).
3) **Transfer** `C:\Archive\Config File Transfer\Civil3D\2023` to the target machine.
4) **Import** on the target machine.
5) **Verify** and complete manual steps.

## Export (source machine)

Run in PowerShell:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -useb https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Export-Civil3D-2023.ps1 | iex"
```

Output bundle:
`C:\Archive\Config File Transfer\Civil3D\2023`

## Import (target machine)

Run in PowerShell:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -useb https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Import-Civil3D-2023.ps1 | iex"
```

## Post-Install Helper (target machine)

Runs import and verify together, with optional interactive prompts to fix remaining items.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -useb https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Post-Install-Civil3D-2023.ps1 | iex"
```

Non-interactive (download to file so you can pass parameters):
```powershell
$script = "$env:TEMP\Post-Install-Civil3D-2023.ps1"
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Post-Install-Civil3D-2023.ps1" -OutFile $script
powershell -ExecutionPolicy Bypass -File $script -Interactive false
```

## Verify (target machine)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -useb https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Verify-Civil3D-2023.ps1 | iex"
```

Non-interactive:
```powershell
$script = "$env:TEMP\Verify-Civil3D-2023.ps1"
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Verify-Civil3D-2023.ps1" -OutFile $script
powershell -ExecutionPolicy Bypass -File $script -Interactive false
```

## Cleanup (optional)

Remove Autodesk installer cache (`C:\Autodesk`) after install:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr -useb https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Remove-Autodesk-InstallerCache.ps1 | iex"
```

## What the scripts do

### Export (`Export-Civil3D-2023.ps1`)
- Exports the current Civil 3D 2023 profile (`HKCU`) to `Civil3D.arg`.
- Copies user Support files from `%APPDATA%\Autodesk\C3D 2023\enu\Support` into the bundle.
- Copies Survey and Pipe Catalog content from ProgramData if present.
- Copies optional Lisp folder.
- Caches network assets (AV Rotate and Copy Shortcut Support File.bat) into `NetworkCache` if `S:` is available.
- Logs to `C:\Archive\Logs\Civil3D\2023`.

### Import (`Import-Civil3D-2023.ps1`)
- Clears the existing profile key, imports `Civil3D.arg`, and sets it as current.
- Scrubs MRU/recent file entries.
- Sets backup options: creates `C:\CADBackup`, sets `AcetMoveBak`, and forces `ISAVEBAK=1`.
- Ensures trusted path includes AV Rotate.
- Copies Support, Survey, Pipe Catalog, and Lisp from the bundle (if paths exist).
- Runs `Copy Shortcut Support File.bat` if found on `S:` (with timeout); otherwise logs and skips.
- Logs to `C:\Archive\Logs\Civil3D\2023`.

### Verify (`Verify-Civil3D-2023.ps1`)
- Validates profile, support paths, plot style path, trusted paths, backup settings.
- Checks ProgramData `enu` path and pipe catalog presence.
- Reports `.dwg` association and installer cache status.
- Summarizes Pass/Warn/Fail with exit codes.
- Interactive prompts can open Default Apps or run cleanup if enabled.

### Post-Install (`Post-Install-Civil3D-2023.ps1`)
- Downloads and runs import + verify in one command.
- Optionally runs cleanup script if present.

### Cleanup (`Remove-Autodesk-InstallerCache.ps1`)
- Safely removes `C:\Autodesk` if present (admin required).

## What is still manual (not automated)
- Uninstall previous C3D versions (Autodesk Uninstall Tool).
- Install C3D 2023 and launch once to initialize ProgramData folders.
- Set `.dwg` default app to AutoCAD DWG Launcher.
- Delete `C:\Autodesk` installer cache (or run cleanup script).
- Run `MOVEBAK C:\CADBackup` in C3D command line (per SOP).
- AV Rotate toolbar UI steps (load `custom_ucs.cuix`, save workspace).
- Any FAQ fixes (Sheet Set Manager reg file, ADFS reg key, etc.).

## Logs
All scripts write logs to: `C:\Archive\Logs\Civil3D\2023`
