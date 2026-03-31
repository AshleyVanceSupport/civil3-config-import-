#Requires -Version 5.1
<#
.SYNOPSIS
    Post-install helper that runs Civil 3D 2023 import and verification in sequence.

.DESCRIPTION
    Downloads Import-Civil3D-2023.ps1 and Verify-Civil3D-2023.ps1 from GitHub,
    runs import, then runs verification. Optionally downloads and passes the cleanup
    script to the verify step. Exits with the highest severity exit code seen.

.PARAMETER BundleRoot
    Path to the Civil 3D 2023 configuration bundle passed to the import script.

.PARAMETER ImportUrl
    URL for Import-Civil3D-2023.ps1 on GitHub.

.PARAMETER VerifyUrl
    URL for Verify-Civil3D-2023.ps1 on GitHub.

.PARAMETER CleanupUrl
    URL for Remove-Autodesk-InstallerCache.ps1 on GitHub.

.PARAMETER LogRoot
    Root directory for log output. A 2023 subfolder is created automatically.

.PARAMETER Interactive
    Set to "true" to enable interactive prompts inside the verify script.

.NOTES
    Run As: User context (not SYSTEM)
    PS Version: 5.1+
    Exit Codes: 0=success, 1=partial/warnings, 2=critical
#>

param(
    [string]$BundleRoot  = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ImportUrl   = "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Import-Civil3D-2023.ps1",
    [string]$VerifyUrl   = "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Verify-Civil3D-2023.ps1",
    [string]$CleanupUrl  = "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Remove-Autodesk-InstallerCache.ps1",
    [string]$LogRoot     = "C:\Archive\Logs\Civil3D",
    [string]$Interactive = "true"
)

# NinjaOne environment variable overrides
if (-not [string]::IsNullOrWhiteSpace($env:BundleRoot))  { $BundleRoot  = $env:BundleRoot }
if (-not [string]::IsNullOrWhiteSpace($env:ImportUrl))   { $ImportUrl   = $env:ImportUrl }
if (-not [string]::IsNullOrWhiteSpace($env:VerifyUrl))   { $VerifyUrl   = $env:VerifyUrl }
if (-not [string]::IsNullOrWhiteSpace($env:CleanupUrl))  { $CleanupUrl  = $env:CleanupUrl }
if (-not [string]::IsNullOrWhiteSpace($env:LogRoot))     { $LogRoot     = $env:LogRoot }
if (-not [string]::IsNullOrWhiteSpace($env:Interactive)) { $Interactive = $env:Interactive }

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
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
}

function Convert-ToBoolean {
    [CmdletBinding()]
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $lower = $Value.ToLower().Trim()
    return ($lower -eq "true" -or $lower -eq "1" -or $lower -eq "yes" -or $lower -eq "on")
}

function Invoke-ScriptDownload {
    [CmdletBinding()]
    param(
        [string]$Url,
        [string]$Destination
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $false
    }

    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -ErrorAction Stop
        return $true
    }
    catch {
        $currentErr = $_
        Write-Warn "Failed to download ${Url}: $($currentErr.Exception.Message)"
        return $false
    }
}

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir  = Join-Path -Path $LogRoot -ChildPath "2023"
$logPath = Join-Path -Path $logDir  -ChildPath "PostInstall_Civil3D_2023_$timestamp.log"

try {
    Initialize-Directory $logDir
    Start-Transcript -Path $logPath -Force -ErrorAction Stop | Out-Null

    Write-Info "User: $env:USERNAME"
    Write-Info "BundleRoot: $BundleRoot"

    $tempRoot   = Join-Path -Path $env:TEMP -ChildPath "Civil3D_2023_PostInstall"
    Initialize-Directory $tempRoot

    $importPath  = Join-Path -Path $tempRoot -ChildPath "Import-Civil3D-2023.ps1"
    $verifyPath  = Join-Path -Path $tempRoot -ChildPath "Verify-Civil3D-2023.ps1"
    $cleanupPath = Join-Path -Path $tempRoot -ChildPath "Remove-Autodesk-InstallerCache.ps1"

    $hasImport  = Invoke-ScriptDownload -Url $ImportUrl  -Destination $importPath
    $hasVerify  = Invoke-ScriptDownload -Url $VerifyUrl  -Destination $verifyPath
    $hasCleanup = Invoke-ScriptDownload -Url $CleanupUrl -Destination $cleanupPath

    if (-not $hasImport) {
        Write-Err "Import script download failed."
        $exitCode = 2
        throw "Import download failed"
    }

    $importExit = 0
    Write-Info "Running import script..."
    & powershell -ExecutionPolicy Bypass -File $importPath -BundleRoot $BundleRoot
    $importExit = $LASTEXITCODE
    Write-Info "Import exit code: $importExit"

    $verifyExit = 1
    if ($hasVerify) {
        Write-Info "Running verification script..."
        if ($hasCleanup) {
            & powershell -ExecutionPolicy Bypass -File $verifyPath -BundleRoot $BundleRoot -Interactive $Interactive -CleanupScriptPath $cleanupPath
        }
        else {
            & powershell -ExecutionPolicy Bypass -File $verifyPath -BundleRoot $BundleRoot -Interactive $Interactive
        }
        $verifyExit = $LASTEXITCODE
        Write-Info "Verify exit code: $verifyExit"
    }
    else {
        Write-Warn "Verify script not available."
    }

    if ($importExit -eq 2 -or $verifyExit -eq 2) {
        $exitCode = 2
    }
    elseif ($importExit -eq 1 -or $verifyExit -eq 1) {
        $exitCode = 1
    }
}
catch {
    $currentErr = $_
    Write-Err "FATAL: $($currentErr.Exception.Message)"
    Write-Err "Stack trace: $($currentErr.ScriptStackTrace)"
    $exitCode = 2
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

exit $exitCode
