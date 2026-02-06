# Civil 3D 2023 export script (run as user)

param(
    [string]$BundleRoot = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ProfileName = "",
    [string]$LogRoot = "C:\Archive\Logs\Civil3D",
    [string]$NetworkCacheRoot = "",
    [string]$ShortcutBatchPath = "S:\Setup Files\CAD\AC3D\Copy Shortcut Support File.bat",
    [string]$AvRotatePath = "S:\Templates\Civil Templates\CAD Tools\AV Rotate"
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

function Test-IsAdmin {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsSystem {
    return [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
}

function Get-CurrentProfileName {
    param([string]$ProfilesRoot)
    try {
        $props = Get-ItemProperty -Path $ProfilesRoot -ErrorAction Stop
        if ($props.CurrentProfile) {
            return $props.CurrentProfile
        }
    }
    catch {
    }
    return $null
}

function Get-FirstProfileName {
    param([string]$ProfilesRoot)
    try {
        $subKey = Get-ChildItem -Path $ProfilesRoot -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $subKey) {
            return $subKey.PSChildName
        }
    }
    catch {
    }
    return $null
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

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $LogRoot "2023"
$logPath = Join-Path $logDir "Export_Civil3D_2023_$timestamp.log"

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

    Ensure-Directory $BundleRoot
    $profileDir = Join-Path $BundleRoot "Profile"
    Ensure-Directory $profileDir

    if ([string]::IsNullOrWhiteSpace($NetworkCacheRoot)) {
        $NetworkCacheRoot = Join-Path $BundleRoot "NetworkCache"
    }
    Ensure-Directory $NetworkCacheRoot

    $profilesRoot = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles"
    if (-not (Test-Path -LiteralPath $profilesRoot)) {
        Write-Err "Civil 3D 2023 profile registry root not found."
        $exitCode = 2
        throw "Profiles registry root missing"
    }

    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        $ProfileName = Get-CurrentProfileName -ProfilesRoot $profilesRoot
    }
    if ([string]::IsNullOrWhiteSpace($ProfileName)) {
        $ProfileName = Get-FirstProfileName -ProfilesRoot $profilesRoot
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

    $profileRegKey = "HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles\$ProfileName"
    $profileRegKeyPs = "Registry::" + $profileRegKey
    if (-not (Test-Path -LiteralPath $profileRegKeyPs)) {
        Write-Err "Profile registry key not found: $ProfileName"
        $exitCode = 2
        throw "Profile registry key missing"
    }

    $argPath = Join-Path $profileDir "Civil3D.arg"
    & reg.exe export "$profileRegKey" "$argPath" /y | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Err "Registry export failed with exit code $LASTEXITCODE"
        $exitCode = 2
        throw "Reg export failed"
    }
    Write-Info "Exported profile to: $argPath"

    $results = @{
        Copied = 0
        Skipped = 0
        Failed = 0
    }
    $failedItems = New-Object System.Collections.ArrayList

    $roamRoot = Join-Path $env:APPDATA "Autodesk\C3D 2023\enu"
    $supportSource = Join-Path $roamRoot "Support"
    $supportDest = Join-Path $BundleRoot "Support"
    Copy-DirectoryContent -Source $supportSource -Destination $supportDest -Label "Support" -Results $results -Failures $failedItems

    $enuProgramData = Join-Path $env:ProgramData "Autodesk\C3D 2023\enu"
    $surveySource = Join-Path $enuProgramData "Survey"
    $surveyDest = Join-Path $BundleRoot "Survey"
    Copy-DirectoryContent -Source $surveySource -Destination $surveyDest -Label "Survey" -Results $results -Failures $failedItems

    $pipeCatalogCandidates = @(Get-ChildItem -Path $enuProgramData -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Pipe*Catalog*" })
    if ($pipeCatalogCandidates.Count -gt 0) {
        $preferred = $pipeCatalogCandidates | Where-Object { $_.Name -eq "Pipe Catalog" } | Select-Object -First 1
        $pipeCatalogSource = $null
        if ($null -ne $preferred) {
            $pipeCatalogSource = $preferred.FullName
        }
        else {
            $pipeCatalogSource = $pipeCatalogCandidates[0].FullName
        }
        $pipeCatalogDest = Join-Path $BundleRoot "PipeCatalog"
        Copy-DirectoryContent -Source $pipeCatalogSource -Destination $pipeCatalogDest -Label "PipeCatalog" -Results $results -Failures $failedItems
    }
    else {
        Write-Warn "Pipe catalog path not found under $enuProgramData"
    }

    $lispSource = Join-Path $supportSource "Lisp"
    $lispDest = Join-Path $BundleRoot "Lisp"
    Copy-DirectoryContent -Source $lispSource -Destination $lispDest -Label "Lisp" -Results $results -Failures $failedItems

    $avRotateCache = Join-Path $NetworkCacheRoot "AV Rotate"
    Copy-DirectoryContent -Source $AvRotatePath -Destination $avRotateCache -Label "NetworkCache-AVRotate" -Results $results -Failures $failedItems

    $shortcutCache = Join-Path $NetworkCacheRoot "Copy Shortcut Support File.bat"
    Copy-FileToCache -Source $ShortcutBatchPath -Destination $shortcutCache -Label "NetworkCache-ShortcutBatch" -Results $results -Failures $failedItems

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
