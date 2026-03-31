#Requires -Version 5.1
<#
.SYNOPSIS
    Verifies Civil 3D 2023 profile and configuration after import.

.DESCRIPTION
    Checks the profile registry key, support paths, plot style path, trusted paths,
    backup settings, ProgramData enu path, pipe catalog, DWG file association, and
    installer cache status. Summarizes Pass/Warn/Fail with matching exit codes.
    Interactive mode can open Default Apps or invoke the cleanup script.

.PARAMETER BundleRoot
    Path to the Civil 3D 2023 configuration bundle. Used to locate Civil3D.arg.

.PARAMETER ProfileName
    Civil 3D profile name. Auto-detected from the current registry profile if omitted.

.PARAMETER LogRoot
    Root directory for log output. A 2023 subfolder is created automatically.

.PARAMETER AvRotatePath
    Expected trusted path for AV Rotate.

.PARAMETER CadBackupPath
    Expected AutoCAD backup directory (AcetMoveBak).

.PARAMETER Interactive
    Set to "true" to enable interactive prompts for remediation steps.

.PARAMETER CleanupScriptPath
    Explicit path to Remove-Autodesk-InstallerCache.ps1. Auto-detected if omitted.

.NOTES
    Run As: User context (not SYSTEM)
    PS Version: 5.1+
    Exit Codes: 0=success, 1=warnings present, 2=failures present
#>

param(
    [string]$BundleRoot        = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ProfileName       = "",
    [string]$LogRoot           = "C:\Archive\Logs\Civil3D",
    [string]$AvRotatePath      = "S:\Templates\Civil Templates\CAD Tools\AV Rotate",
    [string]$CadBackupPath     = "C:\CADBackup",
    [string]$Interactive       = "true",
    [string]$CleanupScriptPath = ""
)

# NinjaOne environment variable overrides
if (-not [string]::IsNullOrWhiteSpace($env:BundleRoot))        { $BundleRoot        = $env:BundleRoot }
if (-not [string]::IsNullOrWhiteSpace($env:ProfileName))       { $ProfileName       = $env:ProfileName }
if (-not [string]::IsNullOrWhiteSpace($env:LogRoot))           { $LogRoot           = $env:LogRoot }
if (-not [string]::IsNullOrWhiteSpace($env:AvRotatePath))      { $AvRotatePath      = $env:AvRotatePath }
if (-not [string]::IsNullOrWhiteSpace($env:CadBackupPath))     { $CadBackupPath     = $env:CadBackupPath }
if (-not [string]::IsNullOrWhiteSpace($env:Interactive))       { $Interactive       = $env:Interactive }
if (-not [string]::IsNullOrWhiteSpace($env:CleanupScriptPath)) { $CleanupScriptPath = $env:CleanupScriptPath }

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

function Test-WarnLike {
    [CmdletBinding()]
    param(
        [hashtable]$Results,
        [string]$Pattern
    )
    return (@($Results.Warn | Where-Object { $_ -like $Pattern })).Count -gt 0
}

function Get-C3dProgramDataEnuPath {
    [CmdletBinding()]
    param([string]$ProgramDataRoot)

    if ([string]::IsNullOrWhiteSpace($ProgramDataRoot)) {
        return $null
    }

    $autodeskRoot = Join-Path -Path $ProgramDataRoot -ChildPath "Autodesk"
    $candidates = @(
        (Join-Path -Path $autodeskRoot -ChildPath "C3D 2023\enu"),
        (Join-Path -Path $autodeskRoot -ChildPath "C3D 2023\R24.2\enu")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    if (Test-Path -LiteralPath $autodeskRoot) {
        $c3dRoot = Get-ChildItem -Path $autodeskRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "C3D 2023*" } |
            Select-Object -First 1

        if ($null -ne $c3dRoot) {
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

    return $null
}

function Add-VerificationResult {
    [CmdletBinding()]
    param(
        [string]$Status,
        [string]$Message,
        [hashtable]$Results
    )

    switch ($Status) {
        "Pass" { $Results.Pass += $Message }
        "Warn" { $Results.Warn += $Message }
        "Fail" { $Results.Fail += $Message }
    }
}

function Test-PathStatus {
    [CmdletBinding()]
    param(
        [string]$Path,
        [string]$Label,
        [hashtable]$Results
    )

    if (Test-Path -LiteralPath $Path) {
        Add-VerificationResult -Status "Pass" -Message "${Label}: present" -Results $Results
        return $true
    }

    Add-VerificationResult -Status "Warn" -Message "${Label}: missing" -Results $Results
    return $false
}

function Test-RegistryContains {
    [CmdletBinding()]
    param(
        [string]$KeyPath,
        [string]$ValueName,
        [string]$ExpectedFragment,
        [string]$Label,
        [hashtable]$Results
    )

    try {
        $value = (Get-ItemProperty -Path $KeyPath -ErrorAction Stop).$ValueName
        if ($null -eq $value) {
            Add-VerificationResult -Status "Warn" -Message "${Label}: ${ValueName} not set" -Results $Results
            return
        }
        if ($value -like "*${ExpectedFragment}*") {
            Add-VerificationResult -Status "Pass" -Message "${Label}: ${ValueName} contains ${ExpectedFragment}" -Results $Results
        }
        else {
            Add-VerificationResult -Status "Warn" -Message "${Label}: ${ValueName} missing ${ExpectedFragment}" -Results $Results
        }
    }
    catch {
        Add-VerificationResult -Status "Warn" -Message "${Label}: unable to read ${ValueName}" -Results $Results
    }
}

function Test-RegistryEquals {
    [CmdletBinding()]
    param(
        [string]$KeyPath,
        [string]$ValueName,
        [string]$Expected,
        [string]$Label,
        [hashtable]$Results
    )

    try {
        $value = (Get-ItemProperty -Path $KeyPath -ErrorAction Stop).$ValueName
        if ($value -eq $Expected) {
            Add-VerificationResult -Status "Pass" -Message "${Label}: ${ValueName}=${Expected}" -Results $Results
        }
        else {
            Add-VerificationResult -Status "Warn" -Message "${Label}: ${ValueName} is ${value}" -Results $Results
        }
    }
    catch {
        Add-VerificationResult -Status "Warn" -Message "${Label}: unable to read ${ValueName}" -Results $Results
    }
}

function Get-DwgAssociation {
    [CmdletBinding()]
    param()

    $assocResult = [ordered]@{
        Match   = $false
        Details = @()
    }

    $patterns = @("AutoCAD", "Acad", "AcLauncher", "DWGLauncher")
    $foundAny  = $false

    function Test-ProgIdValue {
        [CmdletBinding()]
        param(
            [string]$Label,
            [string]$Value
        )
        if ([string]::IsNullOrWhiteSpace($Value)) { return }
        $assocResult.Details += ("${Label}: ${Value}")
        $script:foundAny = $true
        foreach ($pattern in $patterns) {
            if ($Value -match $pattern) {
                $assocResult.Match = $true
                break
            }
        }
    }

    try {
        $userChoiceKey = "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dwg\UserChoice"
        if (Test-Path -LiteralPath $userChoiceKey) {
            $progId = (Get-ItemProperty -Path $userChoiceKey -ErrorAction Stop).ProgId
            Test-ProgIdValue -Label "UserChoice" -Value $progId
        }
    }
    catch { }

    try {
        $hkcuClass = Get-Item -LiteralPath "Registry::HKCU\Software\Classes\.dwg" -ErrorAction Stop
        $progId = $hkcuClass.GetValue("")
        Test-ProgIdValue -Label "HKCU\\Software\\Classes\\.dwg" -Value $progId
    }
    catch { }

    try {
        $hklmClass = Get-Item -LiteralPath "Registry::HKLM\Software\Classes\.dwg" -ErrorAction Stop
        $progId = $hklmClass.GetValue("")
        Test-ProgIdValue -Label "HKLM\\Software\\Classes\\.dwg" -Value $progId
    }
    catch { }

    try {
        $hkcrClass = Get-Item -LiteralPath "Registry::HKCR\.dwg" -ErrorAction Stop
        $progId = $hkcrClass.GetValue("")
        Test-ProgIdValue -Label "HKCR\\.dwg" -Value $progId
    }
    catch { }

    try {
        $openWithKey = "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dwg\OpenWithProgids"
        if (Test-Path -LiteralPath $openWithKey) {
            $progIds = (Get-ItemProperty -Path $openWithKey -ErrorAction Stop).PSObject.Properties |
                Where-Object { $_.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider") } |
                ForEach-Object { $_.Name }
            foreach ($progId in $progIds) {
                Test-ProgIdValue -Label "OpenWithProgids" -Value $progId
            }
        }
    }
    catch { }

    try {
        $assocOutput = & cmd.exe /c "assoc .dwg" 2>$null
        if (-not [string]::IsNullOrWhiteSpace($assocOutput)) {
            $assocValue = ($assocOutput -split "=", 2)[-1]
            Test-ProgIdValue -Label "assoc" -Value $assocValue
        }
    }
    catch { }

    return $assocResult
}

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir  = Join-Path -Path $LogRoot -ChildPath "2023"
$logPath = Join-Path -Path $logDir  -ChildPath "Verify_Civil3D_2023_$timestamp.log"

$interactiveMode = Convert-ToBoolean $Interactive

$verifyResults = @{
    Pass = @()
    Warn = @()
    Fail = @()
}

try {
    Initialize-Directory $logDir
    Start-Transcript -Path $logPath -Force -ErrorAction Stop | Out-Null

    Write-Info "User: $env:USERNAME"
    Write-Info "BundleRoot: $BundleRoot"

    Test-PathStatus -Path $BundleRoot -Label "Bundle root" -Results $verifyResults | Out-Null

    $argPath = Join-Path -Path (Join-Path -Path $BundleRoot -ChildPath "Profile") -ChildPath "Civil3D.arg"
    if (-not (Test-Path -LiteralPath $argPath)) {
        $argPath = Join-Path -Path $BundleRoot -ChildPath "Civil3D.arg"
    }
    Test-PathStatus -Path $argPath -Label "Civil3D.arg" -Results $verifyResults | Out-Null

    $supportPath = Join-Path -Path $env:APPDATA -ChildPath "Autodesk\C3D 2023\enu\Support"
    Test-PathStatus -Path $supportPath -Label "User Support path" -Results $verifyResults | Out-Null

    $enuProgramData = Get-C3dProgramDataEnuPath -ProgramDataRoot $env:ProgramData
    if ($null -eq $enuProgramData) {
        Add-VerificationResult -Status "Warn" -Message "ProgramData enu path not found" -Results $verifyResults
        $autoRoot = Join-Path -Path $env:ProgramData -ChildPath "Autodesk"
        $c3dDirs = @()
        if (Test-Path -LiteralPath $autoRoot) {
            $c3dDirs = @(Get-ChildItem -Path $autoRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "C3D*" })
        }
        if ($c3dDirs.Count -gt 0) {
            Add-VerificationResult -Status "Warn" -Message ("ProgramData C3D folders: " + (($c3dDirs | ForEach-Object { $_.FullName }) -join "; ")) -Results $verifyResults
        }
        else {
            Add-VerificationResult -Status "Warn" -Message "ProgramData C3D folders: none" -Results $verifyResults
        }
    }
    else {
        Add-VerificationResult -Status "Pass" -Message "ProgramData enu path: ${enuProgramData}" -Results $verifyResults
        Test-PathStatus -Path (Join-Path -Path $enuProgramData -ChildPath "Survey") -Label "Survey path" -Results $verifyResults | Out-Null

        $pipeCandidates = @(Get-ChildItem -Path $enuProgramData -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Pipe*Catalog*" })
        if ($pipeCandidates.Count -gt 0) {
            Add-VerificationResult -Status "Pass" -Message "Pipe catalog path: $($pipeCandidates[0].FullName)" -Results $verifyResults
        }
        else {
            Add-VerificationResult -Status "Warn" -Message "Pipe catalog path not found under ${enuProgramData}" -Results $verifyResults
        }
    }

    Test-PathStatus -Path $CadBackupPath -Label "CAD backup folder" -Results $verifyResults | Out-Null

    $profilesRoot = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles"
    if (-not (Test-Path -LiteralPath $profilesRoot)) {
        Add-VerificationResult -Status "Fail" -Message "Profiles registry root missing" -Results $verifyResults
    }
    else {
        $currentProfile = $null
        try {
            $currentProfile = (Get-ItemProperty -Path $profilesRoot -ErrorAction Stop).CurrentProfile
        }
        catch { }

        if ([string]::IsNullOrWhiteSpace($ProfileName)) {
            $ProfileName = $currentProfile
        }
        if ([string]::IsNullOrWhiteSpace($ProfileName)) {
            $profileKey = Get-ChildItem -Path $profilesRoot -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($null -ne $profileKey) {
                $ProfileName = $profileKey.PSChildName
            }
        }

        if ([string]::IsNullOrWhiteSpace($ProfileName)) {
            Add-VerificationResult -Status "Fail" -Message "Profile name not detected" -Results $verifyResults
        }
        else {
            Add-VerificationResult -Status "Pass" -Message "Profile in use: ${ProfileName}" -Results $verifyResults
            $profilePath = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles\${ProfileName}"

            if (Test-Path -LiteralPath $profilePath) {
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "ACAD" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Support"            -Label "Support search path" -Results $verifyResults
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "ACAD" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Lisp"               -Label "Support search path" -Results $verifyResults
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "ACAD" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Keyboard Shortcuts"  -Label "Support search path" -Results $verifyResults
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "PrinterStyleSheetDir" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Plotters\Plot Styles" -Label "Plot styles" -Results $verifyResults

                Test-RegistryEquals -KeyPath "$profilePath\Variables" -ValueName "ISAVEBAK"     -Expected "1"          -Label "Backup option"  -Results $verifyResults
                Test-RegistryContains -KeyPath "$profilePath\Variables" -ValueName "TRUSTEDPATHS" -ExpectedFragment $AvRotatePath -Label "Trusted paths" -Results $verifyResults
            }
            else {
                Add-VerificationResult -Status "Fail" -Message "Profile registry key missing: ${ProfileName}" -Results $verifyResults
            }
        }
    }

    $fixedGeneral = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\FixedProfile\General"
    Test-RegistryEquals -KeyPath $fixedGeneral -ValueName "AcetMoveBak" -Expected $CadBackupPath -Label "Backup path" -Results $verifyResults

    $dwgAssoc = Get-DwgAssociation
    if ($dwgAssoc.Match) {
        if ($dwgAssoc.Details.Count -gt 0) {
            Add-VerificationResult -Status "Pass" -Message ("DWG association: " + ($dwgAssoc.Details -join "; ")) -Results $verifyResults
        }
        else {
            Add-VerificationResult -Status "Pass" -Message "DWG association: AutoCAD detected" -Results $verifyResults
        }
    }
    else {
        if ($dwgAssoc.Details.Count -gt 0) {
            Add-VerificationResult -Status "Warn" -Message ("DWG association: not AutoCAD (" + ($dwgAssoc.Details -join "; ") + ")") -Results $verifyResults
        }
        else {
            Add-VerificationResult -Status "Warn" -Message "DWG association: not set (no ProgId found)" -Results $verifyResults
        }
    }

    $autodeskInstallerCache = "C:\Autodesk"
    if (Test-Path -LiteralPath $autodeskInstallerCache) {
        Add-VerificationResult -Status "Warn" -Message "Installer cache present: ${autodeskInstallerCache}" -Results $verifyResults
    }
    else {
        Add-VerificationResult -Status "Pass" -Message "Installer cache removed" -Results $verifyResults
    }

    Write-Host ""
    Write-Host "Verification summary:" -ForegroundColor Cyan
    Write-Host "  Pass : $($verifyResults.Pass.Count)" -ForegroundColor Green
    Write-Host "  Warn : $($verifyResults.Warn.Count)" -ForegroundColor Yellow
    Write-Host "  Fail : $($verifyResults.Fail.Count)" -ForegroundColor Red

    if ($verifyResults.Pass.Count -gt 0) {
        Write-Host ""
        Write-Host "Pass:" -ForegroundColor Green
        foreach ($item in $verifyResults.Pass) { Write-Host "  - $item" -ForegroundColor Green }
    }
    if ($verifyResults.Warn.Count -gt 0) {
        Write-Host ""
        Write-Host "Warnings:" -ForegroundColor Yellow
        foreach ($item in $verifyResults.Warn) { Write-Host "  - $item" -ForegroundColor Yellow }
    }
    if ($verifyResults.Fail.Count -gt 0) {
        Write-Host ""
        Write-Host "Failures:" -ForegroundColor Red
        foreach ($item in $verifyResults.Fail) { Write-Host "  - $item" -ForegroundColor Red }
    }

    if ($verifyResults.Fail.Count -gt 0) {
        $exitCode = 2
    }
    elseif ($verifyResults.Warn.Count -gt 0) {
        $exitCode = 1
    }

    if ($interactiveMode) {
        if (Test-WarnLike -Results $verifyResults -Pattern "DWG association:*") {
            $answer = Read-Host "Open Default Apps settings to set .dwg now? (Y/N)"
            if ($answer.Trim().ToLower().StartsWith("y")) {
                Start-Process "ms-settings:defaultapps"
                Write-Info "Search .dwg and set AutoCAD DWG Launcher."
            }
        }

        if (Test-WarnLike -Results $verifyResults -Pattern "Installer cache present:*") {
            $localCleanup = $CleanupScriptPath
            if ([string]::IsNullOrWhiteSpace($localCleanup)) {
                if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
                    $candidate1 = Join-Path -Path $PSScriptRoot -ChildPath "Remove-Autodesk-InstallerCache.ps1"
                    if (Test-Path -LiteralPath $candidate1) { $localCleanup = $candidate1 }
                }
                if ([string]::IsNullOrWhiteSpace($localCleanup)) {
                    $candidate2 = "C:\Archive\Remove-Autodesk-InstallerCache.ps1"
                    if (Test-Path -LiteralPath $candidate2) { $localCleanup = $candidate2 }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($localCleanup) -and (Test-Path -LiteralPath $localCleanup)) {
                $answer = Read-Host "Remove C:\Autodesk installer cache now? (Y/N)"
                if ($answer.Trim().ToLower().StartsWith("y")) {
                    Write-Info "Running cleanup script: $localCleanup"
                    & powershell -ExecutionPolicy Bypass -File $localCleanup
                }
            }
            else {
                Write-Warn "Cleanup script not found. Download Remove-Autodesk-InstallerCache.ps1 and run it."
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
