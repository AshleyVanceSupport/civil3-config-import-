#Requires -Version 5.1
<#
.SYNOPSIS
    Exports Civil 3D 2023 profile and configuration files to a bundle directory.

.DESCRIPTION
    Copies the Civil 3D 2023 profile registry export, support files, survey data,
    pipe catalog, LISP files, and network cache files to a staging bundle directory.
    Must run as the target user (not SYSTEM). Civil 3D must be closed.

.NOTES
    Run As: User context (not SYSTEM)
    PS Version: 5.1+
    Exit Codes: 0=success, 1=partial, 2=critical
#>

param(
    [string]$BundleRoot = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ProfileName = "",
    [string]$LogRoot = "C:\Archive\Logs\Civil3D",
    [string]$NetworkCacheRoot = "",
    [string]$ShortcutBatchPath = "S:\Setup Files\CAD\AC3D\Copy Shortcut Support File.bat",
    [string]$AvRotatePath = "S:\Templates\Civil Templates\CAD Tools\AV Rotate"
)

# NinjaOne environment variable overrides
if (-not [string]::IsNullOrWhiteSpace($env:BundleRoot))        { $BundleRoot = $env:BundleRoot }
if (-not [string]::IsNullOrWhiteSpace($env:ProfileName))       { $ProfileName = $env:ProfileName }
if (-not [string]::IsNullOrWhiteSpace($env:LogRoot))           { $LogRoot = $env:LogRoot }
if (-not [string]::IsNullOrWhiteSpace($env:NetworkCacheRoot))  { $NetworkCacheRoot = $env:NetworkCacheRoot }
if (-not [string]::IsNullOrWhiteSpace($env:ShortcutBatchPath)) { $ShortcutBatchPath = $env:ShortcutBatchPath }
if (-not [string]::IsNullOrWhiteSpace($env:AvRotatePath))      { $AvRotatePath = $env:AvRotatePath }

$ErrorActionPreference = "Stop"

function Write-Info {
    [CmdletBinding()]
    param([string]$Message)
    Write-Host $Message -ForegroundColor Gray
}

function Write-Warn {
    [CmdletBinding()]
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Err {
    [CmdletBinding()]
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Initialize-Directory {
    [CmdletBinding()]
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Test-IsAdmin {
    [CmdletBinding()]
    param()
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-IsSystem {
    [CmdletBinding()]
    param()
    return [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
}

function Get-CurrentProfileName {
    [CmdletBinding()]
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
    [CmdletBinding()]
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
    [CmdletBinding()]
    param([string]$Source, [string]$Destination)

    if (Test-Path -LiteralPath $Destination) {
        try {
            $srcHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Source -ErrorAction Stop).Hash
            $dstHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Destination -ErrorAction Stop).Hash
            if ($srcHash -eq $dstHash) {
                return "Skipped"
            }
        }
        catch {
            try {
                $srcInfo = Get-Item -LiteralPath $Source -ErrorAction Stop
                $dstInfo = Get-Item -LiteralPath $Destination -ErrorAction Stop
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
    [CmdletBinding()]
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
        Initialize-Directory $Destination
    }
    catch {
        $currentErr = $_
        $Results.Failed++
        [void]$Failures.Add("${Label}: destination create failed - $($currentErr.Exception.Message)")
        return
    }

    try {
        $sourceRoot = (Resolve-Path -LiteralPath $Source -ErrorAction Stop).Path
    }
    catch {
        $currentErr = $_
        $Results.Failed++
        [void]$Failures.Add("${Label}: source resolve failed - $($currentErr.Exception.Message)")
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
        Initialize-Directory (Split-Path -Path $destFile -Parent)
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
            $currentErr = $_
            $Results.Failed++
            [void]$Failures.Add("${Label}: $relative - $($currentErr.Exception.Message)")
        }
    }
}

function Copy-FileToCache {
    [CmdletBinding()]
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
        Initialize-Directory (Split-Path -Path $Destination -Parent)
        $result = Copy-FileIfDifferent -Source $Source -Destination $Destination
        if ($result -eq "Copied") {
            $Results.Copied++
        }
        else {
            $Results.Skipped++
        }
    }
    catch {
        $currentErr = $_
        $Results.Failed++
        [void]$Failures.Add("${Label}: $($currentErr.Exception.Message)")
    }
}

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $LogRoot "2023"
$logPath = Join-Path $logDir "Export_Civil3D_2023_$timestamp.log"

try {
    Initialize-Directory $logDir
    Start-Transcript -Path $logPath -Force -ErrorAction Stop | Out-Null

    if (Test-IsSystem) {
        Write-Err "This script must run as a user, not SYSTEM."
        $exitCode = 2
        throw "SYSTEM context detected"
    }

    $isAdmin = Test-IsAdmin
    Write-Info "User: $env:USERNAME | Admin: $isAdmin"

    Initialize-Directory $BundleRoot
    $profileDir = Join-Path $BundleRoot "Profile"
    Initialize-Directory $profileDir

    if ([string]::IsNullOrWhiteSpace($NetworkCacheRoot)) {
        $NetworkCacheRoot = Join-Path $BundleRoot "NetworkCache"
    }
    Initialize-Directory $NetworkCacheRoot

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
    $currentErr = $_
    if ($exitCode -eq 0) {
        $exitCode = 2
    }
    Write-Err "FATAL: $($currentErr.Exception.Message)"
    Write-Err "Stack trace: $($currentErr.ScriptStackTrace)"
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

exit $exitCode
