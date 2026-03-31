#Requires -Version 5.1
<#
.SYNOPSIS
    Removes the Autodesk installer cache directory.

.DESCRIPTION
    Safely removes C:\Autodesk (or the specified TargetPath) if it exists.
    Requires administrator privileges. Exits 100 if the cache is already absent.

.PARAMETER TargetPath
    Path to the Autodesk installer cache. Defaults to C:\Autodesk.

.PARAMETER LogRoot
    Root directory for log output. A 2023 subfolder is created automatically.

.NOTES
    Run As: Administrator
    PS Version: 5.1+
    Exit Codes: 0=success, 1=partial, 2=critical, 100=nothing-to-do
#>

param(
    [string]$TargetPath = "C:\Autodesk",
    [string]$LogRoot    = "C:\Archive\Logs\Civil3D"
)

# NinjaOne environment variable overrides
if (-not [string]::IsNullOrWhiteSpace($env:TargetPath)) { $TargetPath = $env:TargetPath }
if (-not [string]::IsNullOrWhiteSpace($env:LogRoot))    { $LogRoot    = $env:LogRoot }

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

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir  = Join-Path -Path $LogRoot -ChildPath "2023"
$logPath = Join-Path -Path $logDir  -ChildPath "Remove_Autodesk_InstallerCache_$timestamp.log"

try {
    Initialize-Directory $logDir
    Start-Transcript -Path $logPath -Force -ErrorAction Stop | Out-Null

    Write-Info "Target: $TargetPath"

    if (-not (Test-Path -LiteralPath $TargetPath)) {
        Write-Info "Nothing to do. Installer cache not found."
        $exitCode = 100
    }
    else {
        $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Err "Administrator privileges required."
            $exitCode = 2
        }
        else {
            try {
                Remove-Item -LiteralPath $TargetPath -Recurse -Force -ErrorAction Stop
                Write-Info "Removed installer cache: $TargetPath"
            }
            catch {
                $currentErr = $_
                Write-Err "Failed to remove installer cache: $($currentErr.Exception.Message)"
                $exitCode = 2
            }
        }
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
