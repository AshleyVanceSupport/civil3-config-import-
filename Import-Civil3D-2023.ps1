# Civil 3D 2023 import script (run as user)

param(
    [string]$BundleRoot = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ProfileArgPath = "",
    [string]$ProfileName = "",
    [string]$LogRoot = "C:\Archive\Logs\Civil3D",
    [string]$NetworkCacheRoot = "",
    [string]$ShortcutBatchPath = "S:\Setup Files\CAD\AC3D\Copy Shortcut Support File.bat",
    [string]$AvRotatePath = "S:\Templates\Civil Templates\CAD Tools\AV Rotate",
    [string]$CadBackupPath = "C:\CADBackup"
)

$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Err {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Ensure-RegistryKey {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

function Get-C3dProgramDataEnuPath {
    param([string]$ProgramDataRoot)

    if ([string]::IsNullOrWhiteSpace($ProgramDataRoot)) {
        return $null
    }

    $autodeskRoot = Join-Path -Path $ProgramDataRoot -ChildPath "Autodesk"
    $candidates = @()
    $candidates += (Join-Path -Path $autodeskRoot -ChildPath "C3D 2023\enu")
    $candidates += (Join-Path -Path $autodeskRoot -ChildPath "C3D 2023\R24.2\enu")

    Write-Info ("ProgramData search root: " + $autodeskRoot)
    Write-Info ("ProgramData candidates: " + ($candidates -join "; "))

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            Write-Info ("ProgramData path found: " + $candidate)
            return $candidate
        }
    }

    if (Test-Path -LiteralPath $autodeskRoot) {
        $c3dRoot = Get-ChildItem -Path $autodeskRoot -Directory -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -like "C3D 2023*"
        } | Select-Object -First 1

        if ($null -ne $c3dRoot) {
            Write-Info ("ProgramData C3D root detected: " + $c3dRoot.FullName)
            $enuPath = Join-Path -Path $c3dRoot.FullName -ChildPath "enu"
            if (Test-Path -LiteralPath $enuPath) {
                return $enuPath
            }

            $r24EnuPath = Join-Path -Path (Join-Path -Path $c3dRoot.FullName -ChildPath "R24.2") -ChildPath "enu"
            if (Test-Path -LiteralPath $r24EnuPath) {
                return $r24EnuPath
            }
        }
    }

    $fallbackEnu = Join-Path -Path $autodeskRoot -ChildPath "C3D 2023"
    if (Test-Path -LiteralPath $fallbackEnu) {
        $enuPath = Join-Path -Path $fallbackEnu -ChildPath "enu"
        if (Test-Path -LiteralPath $enuPath) {
            Write-Info ("ProgramData path found: " + $enuPath)
            return $enuPath
        }
        $r24EnuPath = Join-Path -Path (Join-Path -Path $fallbackEnu -ChildPath "R24.2") -ChildPath "enu"
        if (Test-Path -LiteralPath $r24EnuPath) {
            Write-Info ("ProgramData path found: " + $r24EnuPath)
            return $r24EnuPath
        }
    }

    return $null
}

function Test-IsAdmin {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsSystem {
    return [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
}

function Get-ProfileNameFromArg {
    param([string]$ArgPath)
    $lines = Get-Content -LiteralPath $ArgPath -ErrorAction Stop
    foreach ($line in $lines) {
        if ($line -match "\\Profiles\\<<(.+?)>>") {
            return "<<$($matches[1])>>"
        }
    }
    return $null
}

function Add-SemicolonPath {
    param(
        [string]$Existing,
        [string]$PathToAdd
    )

    if ([string]::IsNullOrWhiteSpace($PathToAdd)) {
        return $Existing
    }

    $normalizedAdd = $PathToAdd.Trim().TrimEnd("\\")
    $entries = @()
    if (-not [string]::IsNullOrWhiteSpace($Existing)) {
        $entries = @($Existing -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    $exists = $false
    foreach ($entry in $entries) {
        if ($entry.Trim().TrimEnd("\\").ToLower() -eq $normalizedAdd.ToLower()) {
            $exists = $true
            break
        }
    }

    if (-not $exists) {
        $entries += $normalizedAdd
    }

    return ($entries -join ";")
}

function Copy-FileIfDifferent {
    param([string]$Source, [string]$Destination)

    if (Test-Path -LiteralPath $Destination) {
        try {
            $srcHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Source).Hash
            $dstHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Destination).Hash
            if ($srcHash -eq $dstHash) {
                return "Skipped"
            }
        }
        catch {
            try {
                $srcInfo = Get-Item -LiteralPath $Source
                $dstInfo = Get-Item -LiteralPath $Destination
                if ($srcInfo.Length -eq $dstInfo.Length -and $srcInfo.LastWriteTimeUtc -eq $dstInfo.LastWriteTimeUtc) {
                    return "Skipped"
                }
            }
            catch {
                return "Copy"
            }
        }
    }

    Copy-Item -LiteralPath $Source -Destination $Destination -Force
    return "Copied"
}

function Copy-DirectoryContent {
    param(
        [string]$Source,
        [string]$Destination,
        [string]$Label,
        [hashtable]$Results,
        [System.Collections.ArrayList]$Failures
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        Write-Info "Skip ${Label}: source missing"
        return
    }

    try {
        Ensure-Directory $Destination
    }
    catch {
        $Results.Failed++
        [void]$Failures.Add("${Label}: destination create failed - $($_.Exception.Message)")
        return
    }

    try {
        $sourceRoot = (Resolve-Path -LiteralPath $Source).Path
    }
    catch {
        $Results.Failed++
        [void]$Failures.Add("${Label}: source resolve failed - $($_.Exception.Message)")
        return
    }

    $files = @(Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -ErrorAction SilentlyContinue)

    if ($files.Count -eq 0) {
        Write-Info "Skip ${Label}: no files"
        return
    }

    foreach ($file in $files) {
        $relative = $file.FullName.Substring($sourceRoot.Length).TrimStart("\\")
        $destFile = Join-Path $Destination $relative
        Ensure-Directory (Split-Path -Path $destFile -Parent)
        try {
            $result = Copy-FileIfDifferent -Source $file.FullName -Destination $destFile
            if ($result -eq "Copied") {
                $Results.Copied++
            }
            else {
                $Results.Skipped++
            }
        }
        catch {
            $Results.Failed++
            [void]$Failures.Add("${Label}: $relative - $($_.Exception.Message)")
        }
    }
}

function Copy-FileToCache {
    param(
        [string]$Source,
        [string]$Destination,
        [string]$Label,
        [hashtable]$Results,
        [System.Collections.ArrayList]$Failures
    )

    if (-not (Test-Path -LiteralPath $Source)) {
        Write-Info "Skip ${Label}: source missing"
        return
    }

    try {
        Ensure-Directory (Split-Path -Path $Destination -Parent)
        $result = Copy-FileIfDifferent -Source $Source -Destination $Destination
        if ($result -eq "Copied") {
            $Results.Copied++
        }
        else {
            $Results.Skipped++
        }
    }
    catch {
        $Results.Failed++
        [void]$Failures.Add("${Label}: $($_.Exception.Message)")
    }
}

function Invoke-ProcessWithTimeout {
    param(
        [string]$FilePath,
        [string]$Arguments,
        [int]$TimeoutMs = 300000
    )

    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $FilePath
    $processInfo.Arguments = $Arguments
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = [System.Diagnostics.Process]::Start($processInfo)
    $completed = $process.WaitForExit($TimeoutMs)
    if (-not $completed) {
        try { $process.Kill() } catch { }
        return @{ TimedOut = $true; ExitCode = $null }
    }

    return @{ TimedOut = $false; ExitCode = $process.ExitCode }
}

function Remove-ProfileMru {
    param([string]$ProfileKeyPath)
    $removed = 0
    $patterns = @(
        "^FileNameMRU\d+$",
        "^MRU$",
        "^Recent.*",
        "^InitialDirectory$"
    )

    if (-not (Test-Path -LiteralPath $ProfileKeyPath)) {
        return 0
    }

    $keys = @($ProfileKeyPath) + @(Get-ChildItem -Path $ProfileKeyPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object { $_.PsPath })
    foreach ($key in $keys) {
        $props = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($null -eq $props) {
            continue
        }
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -in @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider")) {
                continue
            }
            foreach ($pattern in $patterns) {
                if ($prop.Name -match $pattern) {
                    try {
                        Remove-ItemProperty -Path $key -Name $prop.Name -ErrorAction Stop
                        $removed++
                    }
                    catch {
                    }
                    break
                }
            }
        }
    }
    return $removed
}

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $LogRoot "2023"
$logPath = Join-Path $logDir "Import_Civil3D_2023_$timestamp.log"

try {
    Ensure-Directory $logDir
    Start-Transcript -Path $logPath -Force | Out-Null

    if (Test-IsSystem) {
        Write-Err "This script must run as a user, not SYSTEM."
        $exitCode = 2
        throw "SYSTEM context detected"
    }

    $isAdmin = Test-IsAdmin
    Write-Info "User: $env:USERNAME | Admin: $isAdmin"

    if (-not (Test-Path -LiteralPath $BundleRoot)) {
        Write-Err "Bundle root not found: $BundleRoot"
        $exitCode = 2
        throw "Bundle root missing"
    }

    if ([string]::IsNullOrWhiteSpace($NetworkCacheRoot)) {
        $NetworkCacheRoot = Join-Path $BundleRoot "NetworkCache"
    }
    Ensure-Directory $NetworkCacheRoot

    $argCandidates = @()
    $candidate1 = Join-Path (Join-Path $BundleRoot "Profile") "Civil3D.arg"
    $candidate2 = Join-Path $BundleRoot "Civil3D.arg"
    $argCandidates += @($candidate1, $candidate2)

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidate3 = Join-Path $PSScriptRoot "Civil3D.arg"
        $candidate4 = Join-Path (Join-Path $PSScriptRoot "Profile") "Civil3D.arg"
        $argCandidates += @($candidate3, $candidate4)
    }

    if ([string]::IsNullOrWhiteSpace($ProfileArgPath)) {
        foreach ($candidate in $argCandidates) {
            if (Test-Path -LiteralPath $candidate) {
                $ProfileArgPath = $candidate
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($ProfileArgPath)) {
            $argFiles = @(Get-ChildItem -Path $BundleRoot -Filter "*.arg" -File -Recurse -ErrorAction SilentlyContinue)
            if ($argFiles.Count -gt 0) {
                $ProfileArgPath = $argFiles[0].FullName
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($ProfileArgPath) -or -not (Test-Path -LiteralPath $ProfileArgPath)) {
        Write-Err "Civil3D.arg not found."
        Write-Info ("Checked: " + ($argCandidates -join "; "))
        Write-Info "Also searched for *.arg under $BundleRoot"
        Write-Info "You can set -ProfileArgPath explicitly."
        $exitCode = 2
        throw "Profile arg missing"
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        $ProfileName = Get-ProfileNameFromArg -ArgPath $ProfileArgPath
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        Write-Err "Profile name not detected. Use -ProfileName to specify."
        $exitCode = 2
        throw "Profile name missing"
    }

    $acad = Get-Process -Name "acad" -ErrorAction SilentlyContinue
    if ($acad) {
        Write-Err "Civil 3D is running. Close it and re-run this script."
        $exitCode = 1
        throw "Civil 3D running"
    }

    $roamRoot = Join-Path $env:APPDATA "Autodesk\C3D 2023\enu"
    $supportPath = Join-Path $roamRoot "Support"
    Ensure-Directory $supportPath

    $enuProgramData = Get-C3dProgramDataEnuPath -ProgramDataRoot $env:ProgramData
    $surveyPath = $null
    $pipeCatalogPath = $null

    if ($null -eq $enuProgramData) {
        Write-Warn "Civil 3D 2023 ProgramData path not found. Survey and Pipe Catalog copy will be skipped."
        $autoRoot = Join-Path -Path $env:ProgramData -ChildPath "Autodesk"
        $c3dDirs = @()
        if (Test-Path -LiteralPath $autoRoot) {
            $c3dDirs = @(Get-ChildItem -Path $autoRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "C3D*" })
        }
        if ($c3dDirs.Count -gt 0) {
            Write-Info ("ProgramData C3D folders: " + (($c3dDirs | ForEach-Object { $_.FullName }) -join "; "))
        }
        else {
            Write-Info "ProgramData C3D folders: none"
        }
        Write-Warn "If Civil 3D 2023 is installed, launch it once to create ProgramData folders."
        $exitCode = 1
    }
    else {
        $surveyPath = Join-Path $enuProgramData "Survey"

        $pipeCatalogCandidates = @(Get-ChildItem -Path $enuProgramData -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Pipe*Catalog*" })
        if ($pipeCatalogCandidates.Count -gt 0) {
            $preferred = $pipeCatalogCandidates | Where-Object { $_.Name -eq "Pipe Catalog" } | Select-Object -First 1
            if ($null -ne $preferred) {
                $pipeCatalogPath = $preferred.FullName
            }
            else {
                $pipeCatalogPath = $pipeCatalogCandidates[0].FullName
            }
        }
    }

    $backupRoot = Join-Path $logDir "Backup_$timestamp"
    Ensure-Directory $backupRoot

    $profileRegKey = "HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles\$ProfileName"
    $profileRegKeyPs = "Registry::" + $profileRegKey
    if (Test-Path -LiteralPath $profileRegKeyPs) {
        $backupFile = Join-Path $backupRoot "ProfileBackup.reg"
        & reg.exe export "$profileRegKey" "$backupFile" /y | Out-Null
        Remove-Item -LiteralPath $profileRegKeyPs -Recurse -Force -ErrorAction SilentlyContinue
        Write-Info "Backed up and cleared profile: $ProfileName"
    }
    else {
        Write-Info "Profile not found (will be imported): $ProfileName"
    }

    & reg.exe import "$ProfileArgPath" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Registry import failed with exit code $LASTEXITCODE"
        $exitCode = 2
        throw "Reg import failed"
    }

    $profilesRoot = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles"
    if (-not (Test-Path -LiteralPath $profilesRoot)) {
        New-Item -Path $profilesRoot -Force | Out-Null
    }
    New-ItemProperty -Path $profilesRoot -Name "CurrentProfile" -Value $ProfileName -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $profilesRoot -Name "DefaultProfile" -Value $ProfileName -PropertyType String -Force | Out-Null

    $fixedGeneralKey = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\FixedProfile\General"
    Ensure-RegistryKey $fixedGeneralKey
    try {
        Ensure-Directory $CadBackupPath
        New-ItemProperty -Path $fixedGeneralKey -Name "AcetMoveBak" -Value $CadBackupPath -PropertyType String -Force | Out-Null
        Write-Info "Backup folder set: $CadBackupPath"
    }
    catch {
        Write-Warn "Failed to set backup folder: $($_.Exception.Message)"
        $exitCode = 1
    }

    $variablesKey = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles\$ProfileName\Variables"
    Ensure-RegistryKey $variablesKey
    try {
        New-ItemProperty -Path $variablesKey -Name "ISAVEBAK" -Value "1" -PropertyType String -Force | Out-Null
        Write-Info "ISAVEBAK set to 1"
    }
    catch {
        Write-Warn "Failed to set ISAVEBAK: $($_.Exception.Message)"
        $exitCode = 1
    }

    try {
        $currentTrusted = (Get-ItemProperty -Path $variablesKey -ErrorAction SilentlyContinue).TRUSTEDPATHS
        $updatedTrusted = Add-SemicolonPath -Existing $currentTrusted -PathToAdd $AvRotatePath
        if ($updatedTrusted -ne $currentTrusted) {
            New-ItemProperty -Path $variablesKey -Name "TRUSTEDPATHS" -Value $updatedTrusted -PropertyType String -Force | Out-Null
            Write-Info "Trusted path added: $AvRotatePath"
        }
        else {
            Write-Info "Trusted path already present: $AvRotatePath"
        }
    }
    catch {
        Write-Warn "Failed to update trusted paths: $($_.Exception.Message)"
        $exitCode = 1
    }

    $profileKeyPath = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles\$ProfileName"
    $mruRemoved = Remove-ProfileMru -ProfileKeyPath $profileKeyPath
    if ($mruRemoved -gt 0) {
        Write-Info "Removed $mruRemoved MRU entries"
    }

    $results = @{
        Copied = 0
        Skipped = 0
        Failed = 0
    }
    $failedItems = New-Object System.Collections.ArrayList

    $bundleSupport = Join-Path $BundleRoot "Support"
    $bundleSurvey = Join-Path $BundleRoot "Survey"
    $bundlePipe = Join-Path $BundleRoot "PipeCatalog"
    $bundleLisp = Join-Path $BundleRoot "Lisp"

    Copy-DirectoryContent -Source $bundleSupport -Destination $supportPath -Label "Support" -Results $results -Failures $failedItems
    if ($null -ne $surveyPath) {
        Copy-DirectoryContent -Source $bundleSurvey -Destination $surveyPath -Label "Survey" -Results $results -Failures $failedItems
    }
    else {
        Write-Info "Skip Survey: ProgramData path unavailable"
    }

    if ($null -ne $pipeCatalogPath) {
        Copy-DirectoryContent -Source $bundlePipe -Destination $pipeCatalogPath -Label "PipeCatalog" -Results $results -Failures $failedItems
    }
    elseif ($null -ne $enuProgramData) {
        Write-Warn "Pipe catalog path not found under $enuProgramData"
        $exitCode = 1
    }
    else {
        Write-Info "Skip PipeCatalog: ProgramData path unavailable"
    }

    Copy-DirectoryContent -Source $bundleLisp -Destination (Join-Path $supportPath "Lisp") -Label "Lisp" -Results $results -Failures $failedItems

    $avRotateCache = Join-Path $NetworkCacheRoot "AV Rotate"
    Copy-DirectoryContent -Source $AvRotatePath -Destination $avRotateCache -Label "NetworkCache-AVRotate" -Results $results -Failures $failedItems

    $shortcutCache = Join-Path $NetworkCacheRoot "Copy Shortcut Support File.bat"
    Copy-FileToCache -Source $ShortcutBatchPath -Destination $shortcutCache -Label "NetworkCache-ShortcutBatch" -Results $results -Failures $failedItems

    if (Test-Path -LiteralPath $ShortcutBatchPath) {
        Write-Info "Running shortcut support batch: $ShortcutBatchPath"
        $runResult = Invoke-ProcessWithTimeout -FilePath "cmd.exe" -Arguments ("/c `"{0}`"" -f $ShortcutBatchPath) -TimeoutMs 300000
        if ($runResult.TimedOut) {
            Write-Warn "Shortcut support batch timed out"
            $exitCode = 1
        }
        elseif ($runResult.ExitCode -ne 0) {
            Write-Warn "Shortcut support batch exit code: $($runResult.ExitCode)"
            $exitCode = 1
        }
        else {
            Write-Info "Shortcut support batch completed successfully"
        }
    }
    else {
        if (Test-Path -LiteralPath $shortcutCache) {
            Write-Warn "Shortcut support batch not found on S:. Cached copy available at: $shortcutCache"
        }
        else {
            Write-Info "Shortcut support batch not found on S:. Skipping"
        }
    }

    Write-Host ""
    Write-Host "Results:" -ForegroundColor Cyan
    Write-Host "  Copied : $($results.Copied)" -ForegroundColor Green
    Write-Host "  Skipped: $($results.Skipped)" -ForegroundColor Gray
    Write-Host "  Failed : $($results.Failed)" -ForegroundColor Red

    if ($failedItems.Count -gt 0) {
        Write-Warn "Some items failed to copy. Review the log for details."
        $exitCode = 1
    }
}
catch {
    if ($exitCode -eq 0) {
        $exitCode = 2
    }
    Write-Err "FATAL: $($_.Exception.Message)"
    Write-Err "Stack trace: $($_.ScriptStackTrace)"
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

exit $exitCode
