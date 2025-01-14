<#
.SYNOPSIS
    Das script ergänzt die Datei Install.wim durch ein weiteres Image. 
    Fehlt die Datei, wird eine neue erzeugt.
.DESCRIPTION
    Powershell (core) Portierung des c't wimage batch skripts.
    Die Portierung ist NICHT von Heise bzw. c't Magazin
    Im Gegensatz zum Original wird nur 64bit Windows 10 oder neuer unterstuetzt.
    Das script ist für eher fortgeschrittene Anwender gedacht.
    Erstellt: 2025 

    Originalversion (NICHT Powershell-Portierung): Erstellt 2014-2021 von Axel Vahldiek/c't / mailto: axv@ct.de / Version 3.1

    Originalkommentar :
    Bitte lesen Sie unbedingt die Anleitungen zu diesem Skript, siehe https://ct.de/wimage

    Das script versteht folgende Befehlszeilen-Argumente:
    -WhatIf                          - Nur Simulation - keine Änderungen durchführen
    -CleanupOnly                     - Nur cleanup ausführen
    -Shutdown                        - Windows nach Abschluss der Sicherung herunterfahren
    -MaxImageCount                   - Prüfen ob die angegebene Anzahl Images in der Datei Install.wim überschritten ist
    -ImageDescription "Beschreibung" - "Beschreibung" als Beschreibung der Sicherung speichern


    Das script muss mit Administratorrechten gestartet werden!
#>

#Requires -RunAsAdministrator 
#Requires -Version 7.0

# todo: damit geht param nicht mehr => warum? => The term 'param' is not recognized as a name of a cmdlet
# Set-StrictMode -Version 1.0

# ShouldProcess => -WhatIf switch
[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string] $ImageDescription = $null,
    # Maximum allowed images in install.win => 0 means: do not check
    [int] $MaxImageCount = 0,
    # Shutdown computer after script has finished
    [switch] $Shutdown = $false,
    # Only do the cleanup
    [switch] $CleanupOnly = $false
)



function WriteNewline
{
    Write-Host "`n"
}

function WriteProgress
{
    param (
        [string] $progressMessage
    )
    Write-Host $progressMessage -ForegroundColor DarkGreen
}

function WriteExeOutput
{
    param (
        [string] $exeOutput
    )
    Write-Host $exeOutput
}

function WriteWarning
{
    param (
        [string] $text
    )
    #Write-Host ("Warnung: " + $text) -ForegroundColor Yellow
    Write-Warning $text
}

function WriteInfo
{
    param (
        [string] $text
    )
    Write-Host ("Info: " + $text) -ForegroundColor Blue
}


# read shadow ID from the generated CMD script file
function GetLastShadowId
{
    param (
        [Parameter(Mandatory = $true)]
        [string]$shadowInfosFilePath
    )


    if ((Test-Path $shadowInfosFilePath) -ne $true)
    {
        # Shadow ID file not found 
        return $null
    }
    $shadowIdLine = Get-Content -Path $shadowInfosFilePath | Select-String "SET SHADOW_ID_1="
    if ($shadowIdLine)
    {
        return $shadowIdLine.Line.Split('=')[1]
    }
    else
    {
        WriteWarning "Shadow ID not found in the file."
        return $null
    }
}

function IsRecoveryEnabled
{
    [OutputType([Boolean])]
    param (
    )

    $statusPrefix = "Windows RE status:"
    # execute program with parameter(s) and save console output in variable
    $reagentOutput = ReAgentc.exe /info

    $statusLine = $reagentOutput | Select-String $statusPrefix
    $statusText = $statusLine.Line.Split($statusPrefix)[1].Trim()

    if ($statusText -eq "Enabled")
    {
        return $true
    }
    elseif ($statusText -eq "Disabled")
    {
        return $false
    }
    else
    {
        throw "Unexpected ReAgentc status: '$statusText'"
    }
}


function Cleanup
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
    )
    


    WriteNewline
    WriteProgress "*** Aufräumen ***"
    WriteNewline

    WriteProgress "*** Windows RE wieder an alte Stelle zurück verschieben (RE einschalten) ***"
    WriteNewline

    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Enable recovery environment"))
    {
        & reagentc /enable | Out-Null
        if (-not $?)
        {
            WriteWarning "Error in reagentc enable" 
        }
        else
        {
            [Boolean] $recoveryEnvIsEnabledAfterCleanUp = IsRecoveryEnabled
            if (-not $recoveryEnvIsEnabledAfterCleanUp)
            {
                WriteWarning "'reagentc /enable' hat keinen Fehler gemeldet, RE ist trotzdem noch ausgeschaltet!" 
            }
        }
    }

    WriteProgress "*** RunOnce-Schluessel wieder loeschen ***"
    WriteNewline
    # .bat: reg delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce /v "enablewinre" /f >nul 2>nul
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "enablewinre" -Force -ErrorAction SilentlyContinue	



    WriteNewline
    WriteProgress "*** Schattenkopie wieder entfernen ***"
    WriteNewline
    # $shadowId is empty when -whatif/ShouldProcess is active
    if ($PSCmdlet.ShouldProcess("ID of formerly created shadow-copy", "Delete shadow-copy"))
    {
        if ([string]::IsNullOrEmpty($shadowId))
        {
            WriteWarning "shadowId ist leer - löschen übersprungen"
        }
        else
        {
            & "$vshadowBinFile" -ds="$shadowId" | Out-Null
            if (-not $?)
            {
                WriteWarning "Error in vshadow remove"
            }
        }
    }


    WriteProgress "*** Defender Ausnahme fuer $dismBinFile loeschen ***"
    if ($PSCmdlet.ShouldProcess($dismBinFile, "Remove defender exclusion"))
    {
        Remove-MpPreference -ExclusionProcess $dismBinFile
    }
}


# Hint: $? contains last error for powershell commands that used "-ErrorAction SilentlyContinue"
# For native commands (executables started with '&'), $? is set to True when $LASTEXITCODE is 0, and set to False when $LASTEXITCODE is any other value.

# "& some.exe | Out-Null" => "| Out-Null" prevents displaying the console output of the executable



[string] $workdir = $PSScriptRoot
[string] $shadowId = "" # initialize here, so it can be used in Cleanup function

WriteProgress "***************************************"
WriteProgress "*** Willkommen bei wimage_posh v0.1 ***"
WriteProgress "***************************************"
WriteNewline
WriteProgress "*** Einige Pruefungen vorab ... ***"
WriteNewline
WriteProgress "workdir: '$workdir'"


# Windows 11 still reports version 10.x !!!
if ((Get-CimInstance Win32_OperatingSystem).Version.StartsWith("10.") -ne $true)
{
    throw "Skript unterstützt nur Windows 10 oder neuer"
}

# posh -eq ignores case !
if ($workdir.Substring(0, 2) -eq $env:SystemDrive.Substring(0, 2))
{
    throw "Arbeitsverzeichnis darf nicht auf dem Windows-Laufwerk liegen"
}

$sourcesSubFolder = Join-Path $workdir 'sources'
if ((Test-Path $sourcesSubFolder) -ne $true) 
{
    throw "Im Arbeitsverzeichnis muss ein Ordner 'Sources' liegen"
}

if ( (Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux).State -ne "Disabled" )
{
    throw "script funktioniert nicht, wenn eine WSL-1-Distribution installiert ist"
}




if ([string]::IsNullOrEmpty($ImageDescription) -eq $false)
{
    WriteInfo "Benutze ImageDescription: '$ImageDescription'"
}


[string] $dismBinFile = Join-Path "$Env:WinDir" "system32\dism.exe"

WriteProgress "*** Defender Ausnahme fuer $dismBinFile setzen ***"
WriteNewline
if ($PSCmdlet.ShouldProcess($dismBinFile, "Adding defender exclusion"))
{
    Add-MpPreference -ExclusionProcess $dismBinFile
}

$is64bit = [Environment]::Is64BitOperatingSystem
if ($is64bit -ne $true)
{
    throw "Dieses Skript ist nur für 64-Bit-Windows"
}


[string] $vshadowBinFile = Join-Path $workdir "sources\vshadowx64.exe"
if ((Test-Path $vshadowBinFile) -ne $true)
{
    throw "vshadow.exe muss im Sources-Unterordner liegen"
}

[string] $shadowInfosFile = Join-Path $workdir 'sources\vshadowtemp.cmd' 

$leftOverShadowId = GetLastShadowId $shadowInfosFile
if ([string]::IsNullOrEmpty($leftOverShadowId) -eq $false)
{
    WriteProgress "*** Lösche übrig gebliebene Schattenkopie ***"
    & "$vshadowBinFile" -ds="$leftOverShadowId" | Out-Null
    if (-not $?)
    {
        WriteWarning "Kann shadow-copy nicht löschen: '$leftOverShadowId'" 
    }
}


[string] $installWimFilePath = Join-Path $workdir "sources\install.wim"

[string] $dismAction = "append-image"
if (Test-Path $installWimFilePath)
{
    if ($MaxImageCount -gt 0)
    {
        [int] $existingImagesCount = (Get-WindowsImage -ImagePath $installWimFilePath).Length
        if ($existingImagesCount -gt $MaxImageCount)
        {
            throw "In der Install.wim sollen nicht mehr als '$MaxImageCount' Images enthalten sein"
        }
    }
    else
    {
        WriteInfo "MaxImageCount ist 0 => Überspringe Prüfung der Zahl vorhandener Images"
    }
}
else
{
    $dismAction = "capture-image /compress:max"
}


# Sofern vorhanden: ct-WIMage.ini verwenden
[string] $iniFilePath = Join-Path $workdir "sources\ct-WIMage.ini"
if (Test-Path $iniFilePath)
{
    $dismAction += " /ConfigFile:$iniFilePath"
}

# Ab Windows 10 1607 Dism-Option /EA verwenden
$windowsBuildNumber = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentBuild).CurrentBuild
if ([int]$windowsBuildNumber -ge 14393)
{
    $dismAction += " /EA"
}


if ($CleanupOnly)
{
    WriteInfo "Option 'CleanupOnly' ist aktiv!"
    Cleanup
    return
}

if ($Shutdown)
{
    WriteInfo "Option '-Shutdown' ist aktiv - computer wird heruntergefahren, wenn wimage fertig ist!"
}

WriteProgress "*** Keine Probleme gefunden, jetzt geht es los ***"
WriteNewline
WriteProgress "*** Vorbereitungen ... ***"
WriteNewline

# .bat: reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce /f /v "enablewinre" /d "reagentc /enable" >nul 2>nul
WriteProgress "*** RunOnce-Schluessel hinzufuegen zum Restaurieren von WinRE nach Wiederherstellung ***"
WriteNewline
# ShouldPrecess not neccessary because New-ItemProperty handles it internally?
#if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Add RunOnce registry key"))
{
    # this is only done, to have the modification in the shadow copy that is about to be created
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "enablewinre" -Value "reagentc /enable" -Force -ErrorAction SilentlyContinue | Out-Null
    if (-not $?)
    {
        throw "Fehler beim Hinzufügen des RunOnce-Schlüssels"
    }
}


WriteProgress "*** Windows RE auf Windows-Partition verschieben (RE ausschalten) ***"
WriteNewline

[Boolean] $recoveryEnvIsEnabled = IsRecoveryEnabled
if ($recoveryEnvIsEnabled)
{
    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Disable recovery environment"))
    {
        & reagentc /disable | Out-Null
        if (-not $?)
        {
            Cleanup
            throw "Error in reagentc disable => probieren: einmal von hand ausführen: reagentc /disable | reagentc /enable" 
        }
    }
}
else
{
    WriteWarning "RE ausschalten ist schon ausgeschaltet!"
    [string] $reIgnoreResponse = Read-Host "Ignorieren? (j/n)"
    if ($reIgnoreResponse -ne "j")
    {
        return -1
    }
}

WriteNewline
WriteProgress "*** Freien Laufwerksbuchstaben fuer Schattenkopie suchen ***"
WriteNewline

$usedDriveLetters = Get-PSDrive -PSProvider FileSystem | where { ($_.Root.Length -eq 3) -and ($_.Root.EndsWith(":\")) } | select -ExpandProperty Name
$candidateDriveLetters = 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'
$freeDriveLetters = $candidateDriveLetters | where { $_ -notin $usedDriveLetters }
if ($freeDriveLetters.Length -lt 1)
{
    Cleanup
    throw "No available drive letter found"
}
$shadowCopyDriveLetter = $freeDriveLetters[0] + ':'
WriteProgress "Verwende $shadowCopyDriveLetter"
WriteNewline

WriteProgress "*** Schattenkopie der Windows-Partition erzeugen ***"
WriteNewline

if ($PSCmdlet.ShouldProcess($shadowCopyDriveLetter, "Create shadow-copy and mount it as drive"))
{
    # create a persistent shadow copy and generate a CMD script file containing environment variables related to the newly created shadow copies
    & "$vshadowBinFile" -p -script="$shadowInfosFile" $env:SystemDrive | Out-Null
    if (-not $?)
    {
        Cleanup
        throw "Error in vshadow create"
    }


    $shadowId = GetLastShadowId $shadowInfosFile
    if ([string]::IsNullOrEmpty($shadowId) )
    {
        Cleanup
        throw "Error: shadowId is empty"
    }
    else
    {
        WriteInfo "shadowId: '$shadowId'"
    }

    WriteProgress "*** Schattenkopie als Laufwerk sichtbar machen ***"
    WriteNewline
    # Expose (mount) the shadow copy as a drive letter
    & "$vshadowBinFile" -el="$shadowId,$shadowCopyDriveLetter" | Out-Null
    if (-not $?)
    {
        Cleanup
        throw "Error in vshadow mount"
    }
}

WriteNewline
WriteProgress "*** Image erstellen / anhängen ***"

$sysDriveLetter = $env:SystemDrive[0]
[double] $freeSpaceSysDriveGb = [math]::Round((Get-PSDrive $sysDriveLetter).Free / 1GB, 2)
$scratchdir = if ($freeSpaceSysDriveGb -ge 20) { $null } else { "/scratchdir:$workdir" }

WriteNewline
WriteProgress "Es kann ziemlich lange dauern, bis die Fortschrittsanzeige erscheint."
WriteProgress "Nach Erreichen der 100 Prozent kann es wieder dauern, bis es weiter geht."
WriteNewline


$osBuildNumber = [System.Environment]::OSVersion.Version.Build
# this returns e.g. "Windows 10 Enterprise" on Windows 11
# $edition = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).ProductName
$edition = (Get-CimInstance Win32_OperatingSystem).Caption
$datetime = Get-Date -Format "yyyy-MM-dd, HH:mm"
if ([string]::IsNullOrEmpty($ImageDescription) )
{
    $ImageDescription = "$datetime $edition Build $osBuildNumber auf $env:COMPUTERNAME"
}

$imageName = "$(Get-Date -Format 'yyyy-MM-dd HH:mm') $env:COMPUTERNAME $edition"

$dismArgs = @(
    "/$dismAction",
    "/imagefile:$installWimFilePath",
    "/capturedir:$shadowCopyDriveLetter",
    "/name:`"$imageName`"",
    "/description:`"$ImageDescription`"",
    "/checkintegrity",
    "/verify"
)
if ($scratchdir) { $dismArgs += $scratchdir }



if ($PSCmdlet.ShouldProcess($dismArgs, "Call dism.exe"))
{
    $process = Start-Process -FilePath $dismBinFile -ArgumentList $dismArgs -NoNewWindow -PassThru -Wait
    if ($process.ExitCode -ne 0)
    {
        Cleanup
        throw "Beim ausfuehren von DISM ist ein Fehler aufgetreten."
    }
}



WriteNewline
WriteProgress "*** Backup-Liste erzeugen ***"
WriteNewline
# .bat used: %windir%\system32\dism /english /get-wiminfo /wimfile:%workdir%sources\install.wim > %workdir%Backupliste.txt
$backupListFilePath = Join-Path $workdir "Backupliste.txt"
$winImage = Get-WindowsImage -ImagePath $installWimFilePath 
WriteInfo "Datei enthält jetzt '$($winImage.Length)' Images"
$winImage | Out-File -FilePath $backupListFilePath -Encoding utf8
if (-not $?)
{
    WriteWarning "Fehler beim erzeugen der Backup-Liste"
}


Cleanup

WriteNewline
WriteProgress "*** Fertig! ***"


if ($Shutdown)
{
    if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Shutting down computer immediately"))
    {
        & shutdown.exe -s -t 0
    }
} 
