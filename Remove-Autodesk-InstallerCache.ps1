# Removes Autodesk installer cache at C:\Autodesk (run as admin)

param(
    [string]$TargetPath = "C:\Autodesk",
    [string]$LogRoot = "C:\Archive\Logs\Civil3D"
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

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $LogRoot "2023"
$logPath = Join-Path $logDir "Remove_Autodesk_InstallerCache_$timestamp.log"

try {
    Ensure-Directory $logDir
    Start-Transcript -Path $logPath -Force | Out-Null

    Write-Info "Target: $TargetPath"
    if (-not (Test-Path -LiteralPath $TargetPath)) {
        Write-Info "Nothing to do. Installer cache not found."
        exit 100
    }

    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Err "Administrator privileges required."
        exit 2
    }

    try {
        Remove-Item -LiteralPath $TargetPath -Recurse -Force -ErrorAction Stop
        Write-Info "Removed installer cache: $TargetPath"
    }
    catch {
        Write-Err "Failed to remove installer cache: $($_.Exception.Message)"
        $exitCode = 2
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
