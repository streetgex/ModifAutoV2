<#Pré-Build
if /I "$(ConfigurationName)"=="Release" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementVersion.ps1" -Mode Prepare -AssemblyInfo "$(ProjectDir)My Project\AssemblyInfo.vb" -Configuration "$(ConfigurationName)"
if /I "$(ConfigurationName)"=="onServer" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementVersion.ps1" -Mode Prepare -AssemblyInfo "$(ProjectDir)My Project\AssemblyInfo.vb" -Configuration "$(ConfigurationName)"
Post-Build
if /I "$(ConfigurationName)"=="onServer" (
    if exist "\\serv-ad1.igbmc.u-strasbg.fr\c$\Program Files\Script Steph\ModifAuto\$(TargetFileName)" copy /Y "\\serv-ad1.igbmc.u-strasbg.fr\c$\Program Files\Script Steph\ModifAuto\$(TargetFileName)" "\\serv-ad1.igbmc.u-strasbg.fr\c$\Program Files\Script Steph\ModifAuto\$(TargetFileName).%date:~6,4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%%time:~6,2%.bak"
    copy /Y "$(TargetPath)" "\\serv-ad1.igbmc.u-strasbg.fr\c$\Program Files\Script Steph\ModifAuto\$(TargetFileName)"
)
set "IncrementVersion="
if /I "$(ConfigurationName)"=="Release" set "IncrementVersion=1"
if /I "$(ConfigurationName)"=="onServer" set "IncrementVersion=1"
if defined IncrementVersion (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(ProjectDir)IncrementVersion.ps1" -Mode Commit -AssemblyInfo "$(ProjectDir)My Project\AssemblyInfo.vb" -Configuration "$(ConfigurationName)"
    if errorlevel 1 exit /b 1
)

exit /b 0
#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Prepare", "Commit")]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$AssemblyInfo,

    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

# Uniquement en Release ou onServer
if ($Configuration -ne "Release" -and $Configuration -ne "onServer") {
    Write-Host "IncrementVersion : configuration $Configuration, aucune action."
    exit 0
}

$VersionFile = Join-Path $PSScriptRoot "BuildVersion.txt"
$PendingFile = Join-Path $PSScriptRoot "BuildVersion.pending.txt"

if (-not (Test-Path -LiteralPath $AssemblyInfo)) {
    throw "Fichier AssemblyInfo introuvable : $AssemblyInfo"
}

# Ligne active uniquement.
# Les lignes commentees sont ignorees.
$AssemblyVersionPattern = '(?m)^(?<prefix>[ \t]*<Assembly:[ \t]*AssemblyVersion\(")(?<version>\d+\.\d+\.\d+\.\d+)(?<suffix>"\)>[ \t]*\r?)$'

$AssemblyFileVersionPattern = '(?m)^(?<prefix>[ \t]*<Assembly:[ \t]*AssemblyFileVersion\(")(?<version>\d+\.\d+\.\d+\.\d+)(?<suffix>"\)>[ \t]*\r?)$'


function Get-AssemblyInfoContent {

    $Content = [System.IO.File]::ReadAllText($AssemblyInfo)

    $Matches = [regex]::Matches(
        $Content,
        $AssemblyVersionPattern
    )

    if ($Matches.Count -eq 0) {
        throw "AssemblyVersion active introuvable dans $AssemblyInfo"
    }

    if ($Matches.Count -gt 1) {
        throw "Plusieurs AssemblyVersion actives trouvees dans $AssemblyInfo"
    }

    return $Content
}


function Get-CurrentAssemblyVersion {

    param (
        [string]$Content
    )

    $Match = [regex]::Match(
        $Content,
        $AssemblyVersionPattern
    )

    return [version]$Match.Groups["version"].Value
}


function Set-AssemblyVersions {

    param (
        [string]$Content,
        [string]$Version
    )

    # AssemblyVersion
    $Content = [regex]::Replace(
        $Content,
        $AssemblyVersionPattern,
        '${prefix}' + $Version + '${suffix}',
        1
    )

    # AssemblyFileVersion si presente
    if ([regex]::IsMatch(
        $Content,
        $AssemblyFileVersionPattern
    )) {

        $Content = [regex]::Replace(
            $Content,
            $AssemblyFileVersionPattern,
            '${prefix}' + $Version + '${suffix}',
            1
        )
    }

    $Encoding = New-Object System.Text.UTF8Encoding($true)

    [System.IO.File]::WriteAllText(
        $AssemblyInfo,
        $Content,
        $Encoding
    )
}


# ============================================================
# PREPARE
# ============================================================

if ($Mode -eq "Prepare") {

    $Content = Get-AssemblyInfoContent
    $CurrentVersion = Get-CurrentAssemblyVersion $Content


    # Premiere utilisation :
    # la version actuelle sert de reference initiale.
    if (-not (Test-Path -LiteralPath $VersionFile)) {

        $CommittedVersion = $CurrentVersion

        Set-Content `
            -LiteralPath $VersionFile `
            -Value $CommittedVersion.ToString() `
            -Encoding ASCII

        Write-Host "Version initiale : $CommittedVersion"
    }
    else {

        $CommittedText = (
            Get-Content `
                -LiteralPath $VersionFile `
                -Raw
        ).Trim()

        try {
            $CommittedVersion = [version]$CommittedText
        }
        catch {
            throw "Version invalide dans $VersionFile : $CommittedText"
        }
    }


    # ========================================================
    # Determination de la prochaine version
    # ========================================================

    # Major modifie manuellement
    if ($CurrentVersion.Major -ne $CommittedVersion.Major) {

        $NextVersion = "{0}.0.0.0" -f `
            $CurrentVersion.Major

        Write-Host "Changement manuel du Major detecte."
        Write-Host "Reset Minor, Build et Revision."
    }


    # Minor (x) modifie manuellement
    elseif ($CurrentVersion.Minor -ne $CommittedVersion.Minor) {

        $NextVersion = "{0}.{1}.0.0" -f `
            $CurrentVersion.Major,
            $CurrentVersion.Minor

        Write-Host "Changement manuel du Minor detecte."
        Write-Host "Reset Build et Revision."
    }


    # Build (y) modifie manuellement
    elseif ($CurrentVersion.Build -ne $CommittedVersion.Build) {

        $NextVersion = "{0}.{1}.{2}.0" -f `
            $CurrentVersion.Major,
            $CurrentVersion.Minor,
            $CurrentVersion.Build

        Write-Host "Changement manuel du Build detecte."
        Write-Host "Reset Revision."
    }


    # Aucun changement manuel de Major / Minor / Build
    # => increment automatique de z
    else {

        $NextRevision = $CommittedVersion.Revision + 1

        if ($NextRevision -gt 65534) {
            throw "La revision maximale 65534 est atteinte."
        }

        $NextVersion = "{0}.{1}.{2}.{3}" -f `
            $CommittedVersion.Major,
            $CommittedVersion.Minor,
            $CommittedVersion.Build,
            $NextRevision
    }


    # Injection avant compilation
    Set-AssemblyVersions `
        -Content $Content `
        -Version $NextVersion


    # Version en attente
    Set-Content `
        -LiteralPath $PendingFile `
        -Value $NextVersion `
        -Encoding ASCII


    Write-Host "========================================"
    Write-Host "Version validee  : $CommittedVersion"
    Write-Host "Version detectee : $CurrentVersion"
    Write-Host "Version preparee : $NextVersion"
    Write-Host "========================================"

    exit 0
}


# ============================================================
# COMMIT
# ============================================================

if ($Mode -eq "Commit") {

    if (-not (Test-Path -LiteralPath $PendingFile)) {
        throw "Aucune version en attente de validation."
    }


    $PendingVersion = (
        Get-Content `
            -LiteralPath $PendingFile `
            -Raw
    ).Trim()


    try {
        [void][version]$PendingVersion
    }
    catch {
        throw "Version invalide dans $PendingFile : $PendingVersion"
    }


    $Content = Get-AssemblyInfoContent
    $CurrentVersion = Get-CurrentAssemblyVersion $Content


    if ($CurrentVersion.ToString() -ne $PendingVersion) {

        throw "AssemblyInfo contient $CurrentVersion alors que la version en attente est $PendingVersion"
    }


    # Le build a reussi :
    # la nouvelle version devient la reference.
    Set-Content `
        -LiteralPath $VersionFile `
        -Value $PendingVersion `
        -Encoding ASCII


    Remove-Item `
        -LiteralPath $PendingFile `
        -Force


    Write-Host "========================================"
    Write-Host "Version validee : $PendingVersion"
    Write-Host "========================================"

    exit 0
}