# Civil 3D 2023 Transfer Notes

## Status Snapshot (Feb 06, 2026)
- Export script ran successfully on source; bundle created at `C:\Archive\Config File Transfer\Civil3D\2023`.
- Import script ran successfully on target; profile applied, MRUs cleared, support files copied.
- ProgramData `enu` path was missing before install; after install, import succeeded.
- `.dwg` association still unset; `C:\Autodesk` installer cache still present.

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

## What is still manual (not automated)
- Uninstall previous C3D versions (Autodesk Uninstall Tool).
- Install C3D 2023 and launch once to initialize ProgramData folders.
- Set `.dwg` default app to AutoCAD DWG Launcher.
- Delete `C:\Autodesk` installer cache (after install).
- Run `MOVEBAK C:\CADBackup` in C3D command line (per SOP).
- AV Rotate toolbar UI steps (load `custom_ucs.cuix`, save workspace).
- Any FAQ fixes (Sheet Set Manager reg file, ADFS reg key, etc.).

## Improvement Plan
- Auto-detect ProgramData `enu` path more robustly (including `C:\ProgramData\Autodesk\C3D 2023`).
- Add optional switch to pull missing ProgramData files from a known network share if available.
- Provide interactive verification to guide remaining manual steps.
- Add optional cleanup script to remove `C:\Autodesk` cache safely.
- Add a one-command wrapper that runs Export -> Copy -> Import -> Verify with logging.

## Current Known Gaps
- Target machine ProgramData `C3D 2023` folder missing; Survey/Pipe Catalog were not copied.
- `S:` availability affects network cache and batch execution.
