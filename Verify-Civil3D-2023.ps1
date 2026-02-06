# Civil 3D 2023 verification script (run as user)

param(
    [string]$BundleRoot = "C:\Archive\Config File Transfer\Civil3D\2023",
    [string]$ProfileName = "",
    [string]$LogRoot = "C:\Archive\Logs\Civil3D",
    [string]$AvRotatePath = "S:\Templates\Civil Templates\CAD Tools\AV Rotate",
    [string]$CadBackupPath = "C:\CADBackup",
    [string]$Interactive = "true",
    [string]$CleanupScriptPath = ""
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

function Has-WarnLike {
    param(
        [hashtable]$Results,
        [string]$Pattern
    )
    return (@($Results.Warn | Where-Object { $_ -like $Pattern })).Count -gt 0
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

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    if (Test-Path -LiteralPath $autodeskRoot) {
        $c3dRoot = Get-ChildItem -Path $autodeskRoot -Directory -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -like "C3D 2023*"
        } | Select-Object -First 1

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

function Add-Result {
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
    param(
        [string]$Path,
        [string]$Label,
        [hashtable]$Results
    )

    if (Test-Path -LiteralPath $Path) {
        Add-Result -Status "Pass" -Message "${Label}: present" -Results $Results
        return $true
    }

    Add-Result -Status "Warn" -Message "${Label}: missing" -Results $Results
    return $false
}

function Test-RegistryContains {
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
            Add-Result -Status "Warn" -Message "${Label}: ${ValueName} not set" -Results $Results
            return
        }
        if ($value -like "*${ExpectedFragment}*") {
            Add-Result -Status "Pass" -Message "${Label}: ${ValueName} contains ${ExpectedFragment}" -Results $Results
        }
        else {
            Add-Result -Status "Warn" -Message "${Label}: ${ValueName} missing ${ExpectedFragment}" -Results $Results
        }
    }
    catch {
        Add-Result -Status "Warn" -Message "${Label}: unable to read ${ValueName}" -Results $Results
    }
}

function Test-RegistryEquals {
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
            Add-Result -Status "Pass" -Message "${Label}: ${ValueName}=${Expected}" -Results $Results
        }
        else {
            Add-Result -Status "Warn" -Message "${Label}: ${ValueName} is ${value}" -Results $Results
        }
    }
    catch {
        Add-Result -Status "Warn" -Message "${Label}: unable to read ${ValueName}" -Results $Results
    }
}

function Get-DwgAssociation {
    $result = [ordered]@{
        Match = $false
        Details = @()
    }

    $patterns = @(
        "AutoCAD",
        "AcLauncher",
        "DWGLauncher"
    )

    function Test-ProgIdValue {
        param(
            [string]$Label,
            [string]$Value
        )
        if ([string]::IsNullOrWhiteSpace($Value)) { return }
        $result.Details += ("${Label}: ${Value}")
        foreach ($pattern in $patterns) {
            if ($Value -match $pattern) {
                $result.Match = $true
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
    catch {
    }

    try {
        $hkcuClass = Get-Item -LiteralPath "Registry::HKCU\Software\Classes\.dwg" -ErrorAction Stop
        $progId = $hkcuClass.GetValue("")
        Test-ProgIdValue -Label "HKCU\\Software\\Classes\\.dwg" -Value $progId
    }
    catch {
    }

    try {
        $hklmClass = Get-Item -LiteralPath "Registry::HKLM\Software\Classes\.dwg" -ErrorAction Stop
        $progId = $hklmClass.GetValue("")
        Test-ProgIdValue -Label "HKLM\\Software\\Classes\\.dwg" -Value $progId
    }
    catch {
    }

    try {
        $hkcrClass = Get-Item -LiteralPath "Registry::HKCR\.dwg" -ErrorAction Stop
        $progId = $hkcrClass.GetValue("")
        Test-ProgIdValue -Label "HKCR\\.dwg" -Value $progId
    }
    catch {
    }

    try {
        $openWithKey = "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.dwg\OpenWithProgids"
        if (Test-Path -LiteralPath $openWithKey) {
            $progIds = (Get-ItemProperty -Path $openWithKey -ErrorAction Stop).PSObject.Properties | Where-Object {
                $_.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider")
            } | ForEach-Object { $_.Name }
            foreach ($progId in $progIds) {
                Test-ProgIdValue -Label "OpenWithProgids" -Value $progId
            }
        }
    }
    catch {
    }

    return $result
}

$exitCode = 0
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logDir = Join-Path $LogRoot "2023"
$logPath = Join-Path $logDir "Verify_Civil3D_2023_$timestamp.log"

$interactiveMode = Convert-ToBoolean $Interactive

$results = @{
    Pass = @()
    Warn = @()
    Fail = @()
}

try {
    Ensure-Directory $logDir
    Start-Transcript -Path $logPath -Force | Out-Null

    Write-Info "User: $env:USERNAME"
    Write-Info "BundleRoot: $BundleRoot"

    Test-PathStatus -Path $BundleRoot -Label "Bundle root" -Results $results | Out-Null

    $argPath = Join-Path (Join-Path $BundleRoot "Profile") "Civil3D.arg"
    if (-not (Test-Path -LiteralPath $argPath)) {
        $argPath = Join-Path $BundleRoot "Civil3D.arg"
    }
    Test-PathStatus -Path $argPath -Label "Civil3D.arg" -Results $results | Out-Null

    $supportPath = Join-Path $env:APPDATA "Autodesk\C3D 2023\enu\Support"
    Test-PathStatus -Path $supportPath -Label "User Support path" -Results $results | Out-Null

    $enuProgramData = Get-C3dProgramDataEnuPath -ProgramDataRoot $env:ProgramData
    if ($null -eq $enuProgramData) {
        Add-Result -Status "Warn" -Message "ProgramData enu path not found" -Results $results
        $autoRoot = Join-Path -Path $env:ProgramData -ChildPath "Autodesk"
        $c3dDirs = @()
        if (Test-Path -LiteralPath $autoRoot) {
            $c3dDirs = @(Get-ChildItem -Path $autoRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "C3D*" })
        }
        if ($c3dDirs.Count -gt 0) {
            Add-Result -Status "Warn" -Message ("ProgramData C3D folders: " + (($c3dDirs | ForEach-Object { $_.FullName }) -join "; ")) -Results $results
        }
        else {
            Add-Result -Status "Warn" -Message "ProgramData C3D folders: none" -Results $results
        }
    }
    else {
        Add-Result -Status "Pass" -Message "ProgramData enu path: ${enuProgramData}" -Results $results
        Test-PathStatus -Path (Join-Path $enuProgramData "Survey") -Label "Survey path" -Results $results | Out-Null

        $pipeCandidates = @(Get-ChildItem -Path $enuProgramData -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*Pipe*Catalog*" })
        if ($pipeCandidates.Count -gt 0) {
            Add-Result -Status "Pass" -Message "Pipe catalog path: $($pipeCandidates[0].FullName)" -Results $results
        }
        else {
            Add-Result -Status "Warn" -Message "Pipe catalog path not found under ${enuProgramData}" -Results $results
        }
    }

    Test-PathStatus -Path $CadBackupPath -Label "CAD backup folder" -Results $results | Out-Null

    $profilesRoot = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles"
    if (-not (Test-Path -LiteralPath $profilesRoot)) {
        Add-Result -Status "Fail" -Message "Profiles registry root missing" -Results $results
    }
    else {
        $currentProfile = $null
        try {
            $currentProfile = (Get-ItemProperty -Path $profilesRoot -ErrorAction Stop).CurrentProfile
        }
        catch {
        }

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
            Add-Result -Status "Fail" -Message "Profile name not detected" -Results $results
        }
        else {
            Add-Result -Status "Pass" -Message "Profile in use: ${ProfileName}" -Results $results
            $profilePath = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\Profiles\${ProfileName}"

            if (Test-Path -LiteralPath $profilePath) {
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "ACAD" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Support" -Label "Support search path" -Results $results
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "ACAD" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Lisp" -Label "Support search path" -Results $results
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "ACAD" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Keyboard Shortcuts" -Label "Support search path" -Results $results
                Test-RegistryContains -KeyPath "$profilePath\General" -ValueName "PrinterStyleSheetDir" -ExpectedFragment "S:\Templates\Civil Templates\REFERENCE\Plotters\Plot Styles" -Label "Plot styles" -Results $results

                Test-RegistryEquals -KeyPath "$profilePath\Variables" -ValueName "ISAVEBAK" -Expected "1" -Label "Backup option" -Results $results
                Test-RegistryContains -KeyPath "$profilePath\Variables" -ValueName "TRUSTEDPATHS" -ExpectedFragment $AvRotatePath -Label "Trusted paths" -Results $results
            }
            else {
                Add-Result -Status "Fail" -Message "Profile registry key missing: ${ProfileName}" -Results $results
            }
        }
    }

    $fixedGeneral = "Registry::HKCU\Software\Autodesk\AutoCAD\R24.2\ACAD-6100:409\FixedProfile\General"
    Test-RegistryEquals -KeyPath $fixedGeneral -ValueName "AcetMoveBak" -Expected $CadBackupPath -Label "Backup path" -Results $results

    $dwgAssoc = Get-DwgAssociation
    if ($dwgAssoc.Match) {
        if ($dwgAssoc.Details.Count -gt 0) {
            Add-Result -Status "Pass" -Message ("DWG association: " + ($dwgAssoc.Details -join "; ")) -Results $results
        }
        else {
            Add-Result -Status "Pass" -Message "DWG association: AutoCAD detected" -Results $results
        }
    }
    else {
        if ($dwgAssoc.Details.Count -gt 0) {
            Add-Result -Status "Warn" -Message ("DWG association: not AutoCAD (" + ($dwgAssoc.Details -join "; ") + ")") -Results $results
        }
        else {
            Add-Result -Status "Warn" -Message "DWG association: not set" -Results $results
        }
    }

    $autodeskInstallerCache = "C:\Autodesk"
    if (Test-Path -LiteralPath $autodeskInstallerCache) {
        Add-Result -Status "Warn" -Message "Installer cache present: ${autodeskInstallerCache}" -Results $results
    }
    else {
        Add-Result -Status "Pass" -Message "Installer cache removed" -Results $results
    }

    Write-Host ""
    Write-Host "Verification summary:" -ForegroundColor Cyan
    Write-Host "  Pass : $($results.Pass.Count)" -ForegroundColor Green
    Write-Host "  Warn : $($results.Warn.Count)" -ForegroundColor Yellow
    Write-Host "  Fail : $($results.Fail.Count)" -ForegroundColor Red

    if ($results.Pass.Count -gt 0) {
        Write-Host ""
        Write-Host "Pass:" -ForegroundColor Green
        foreach ($item in $results.Pass) { Write-Host "  - $item" -ForegroundColor Green }
    }
    if ($results.Warn.Count -gt 0) {
        Write-Host ""
        Write-Host "Warnings:" -ForegroundColor Yellow
        foreach ($item in $results.Warn) { Write-Host "  - $item" -ForegroundColor Yellow }
    }
    if ($results.Fail.Count -gt 0) {
        Write-Host ""
        Write-Host "Failures:" -ForegroundColor Red
        foreach ($item in $results.Fail) { Write-Host "  - $item" -ForegroundColor Red }
    }

    if ($results.Fail.Count -gt 0) {
        $exitCode = 2
    }
    elseif ($results.Warn.Count -gt 0) {
        $exitCode = 1
    }

    if ($interactiveMode) {
        if (Has-WarnLike -Results $results -Pattern "DWG association:*") {
            $answer = Read-Host "Open Default Apps settings to set .dwg now? (Y/N)"
            if ($answer.Trim().ToLower().StartsWith("y")) {
                Start-Process "ms-settings:defaultapps"
                Write-Info "Search .dwg and set AutoCAD DWG Launcher."
            }
        }

        if (Has-WarnLike -Results $results -Pattern "Installer cache present:*") {
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
    Write-Err "FATAL: $($_.Exception.Message)"
    Write-Err "Stack trace: $($_.ScriptStackTrace)"
    $exitCode = 2
}
finally {
    Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
}

exit $exitCode
