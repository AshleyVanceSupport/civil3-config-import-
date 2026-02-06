# Civil 3D 2023 post-install helper (run as user)

param(
    [string]$BundleRoot = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ImportUrl = "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Import-Civil3D-2023.ps1",
    [string]$VerifyUrl = "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Verify-Civil3D-2023.ps1",
    [string]$CleanupUrl = "https://raw.githubusercontent.com/AshleyVanceSupport/civil3-config-import-/refs/heads/main/Remove-Autodesk-InstallerCache.ps1",
    [string]$LogRoot = "C:\Archive\Logs\Civil3D",
    [string]$Interactive = "true"
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

function Convert-ToBoolean {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $lower = $Value.ToLower().Trim()
    return ($lower -eq "true" -or $lower -eq "1" -or $lower -eq "yes" -or $lower -eq "on")
}

function Download-Script {
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
        Write-Warn "Failed to download $Url: $($_.Exception.Message)"
        return $false
    }
}

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $LogRoot "2023"
$logPath = Join-Path $logDir "PostInstall_Civil3D_2023_$timestamp.log"
$interactiveMode = Convert-ToBoolean $Interactive

try {
    Ensure-Directory $logDir
    Start-Transcript -Path $logPath -Force | Out-Null

    Write-Info "User: $env:USERNAME"
    Write-Info "BundleRoot: $BundleRoot"

    $tempRoot = Join-Path $env:TEMP "Civil3D_2023_PostInstall"
    Ensure-Directory $tempRoot

    $importPath = Join-Path $tempRoot "Import-Civil3D-2023.ps1"
    $verifyPath = Join-Path $tempRoot "Verify-Civil3D-2023.ps1"
    $cleanupPath = Join-Path $tempRoot "Remove-Autodesk-InstallerCache.ps1"

    $hasImport = Download-Script -Url $ImportUrl -Destination $importPath
    $hasVerify = Download-Script -Url $VerifyUrl -Destination $verifyPath
    $hasCleanup = Download-Script -Url $CleanupUrl -Destination $cleanupPath

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
        $verifyExit = 1
    }

    if ($importExit -eq 2 -or $verifyExit -eq 2) {
        $exitCode = 2
    }
    elseif ($importExit -eq 1 -or $verifyExit -eq 1) {
        $exitCode = 1
    }
}
catch {
    Write-Err "FATAL: $($_.Exception.Message)"
    Write-Err "Stack trace: $($_.ScriptStackTrace)"
    $exitCode = 2
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

exit $exitCode
