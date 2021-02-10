# Script to generate a report on user's software setup.

# Start stopwatch
$sw = [Diagnostics.Stopwatch]::StartNew()

# Read from prefs file
$prefs = Get-Content .\preferences.json | ConvertFrom-Json
$ServerURL = $prefs.ServerURL

$report = New-Object -TypeName psobject
$dateformat = "yyyy-MM-dd HH:mm"

$report | Add-Member -MemberType NoteProperty -Name ReportServer -Value $ServerURL

$username = $env:UserName
$report | Add-Member -MemberType NoteProperty -Name UserName -Value $username

$computername = $env:ComputerName
$report | Add-Member -MemberType NoteProperty -Name ComputerName -Value $computername

$reportdate = Get-Date -Format $dateformat
$report | Add-Member -MemberType NoteProperty -Name ReportDate -Value $reportdate

function CheckFileInfo {
    param (
        [string]$FilePath
    )
    $settingsinfo = New-Object -TypeName psobject
    $lastwritetimestring = (Get-ChildItem $FilePath).LastWriteTime.toString()
    $lastwritetime = Get-Date $lastwritetimestring -Format $dateformat
    $size = Get-ChildItem $FilePath | Select-Object Length
    $size = [math]::Round(($size.Length/1024),2)
    [string]$sizestr = $size
    $sizestr = $sizestr + " Kb"
    $settingsinfo | Add-Member -MemberType NoteProperty -Name LastWriteTime -Value $lastwritetime
    $settingsinfo | Add-Member -MemberType NoteProperty -Name Size -Value $sizestr
    return $settingsinfo
}
$wwsettingsobj = CheckFileInfo -FilePath "c:\ProgramData\Woodwing\WWsettings.xml"
$report | Add-Member -MemberType NoteProperty -Name WWsettings -Value $wwsettingsobj

$appmappingobj = CheckFileInfo -FilePath "c:\ProgramData\Woodwing\ApplicationMapping.xml"
$report | Add-Member -MemberType NoteProperty -Name ApplicationMapping -Value $appmappingobj

function CheckContentStationAIR {
    [xml]$appinfoxml = Get-Content "C:\Program Files (x86)\Content Station\META-INF\AIR\application.xml"
    $appversionstring = $appinfoxml.application.VersionLabel
    return $appversionstring
}
$csairversion = CheckContentStationAIR
$report | Add-Member -MemberType NoteProperty -Name ContentStationAIRVersion -Value $csairversion

function CheckRegistryForAdobeSoftware {
    $adobesoftware = Get-WmiObject -Class Win32_Product | Where-Object vendor -eq "Adobe Systems Incorporated" | Select-Object Name
    return $adobesoftware
}
$adobesoftwarereg = CheckRegistryForAdobeSoftware
$report | Add-Member -MemberType NoteProperty -Name AdobeSoftwareInRegistry -Value $adobesoftwarereg

function CheckDriveForAdobeSoftware {
    $adobeapps = @(
        "InDesign",
        "InCopy",
        "Photoshop"
    )
    $AdobeExeFiles = New-Object -TypeName psobject
    foreach ($app in $adobeapps){
        $appname = $app + ".exe"
        $appname = Get-Childitem -Path 'C:\Program Files\Adobe\' -Include $appname -File -Recurse -ErrorAction SilentlyContinue | Select-Object FullName
        $appverinfo = Get-Childitem -Path 'C:\Program Files\Adobe\' -Include $appname -File -Recurse -ErrorAction SilentlyContinue | Select-Object VersionInfo | Select-Object FileVersionRaw
        $appinfo = New-Object -TypeName psobject
        $appinfo | Add-Member -MemberType NoteProperty -Name AppName -Value $appname
        $appinfo | Add-Member -MemberType NoteProperty -Name AppVersion -Value $appverinfo
        $AdobeExeFiles | Add-Member -MemberType NoteProperty -Name $app -Value $appinfo
    }
    return $AdobeExeFiles
}
$adobesoftwareondrive = CheckDriveForAdobeSoftware
$report | Add-Member -MemberType NoteProperty -Name AdobeSoftwareOnDrive -Value $adobesoftwareondrive

function CheckMappings {
    $mappings = Get-SMBmapping | Select-Object Status, LocalPath, RemotePath
    return $mappings
}
$mappings = CheckMappings
$report | Add-Member -MemberType NoteProperty -Name Mappings -Value $mappings


$reportjson = ConvertTo-Json $report
Out-File -FilePath .\report.json -InputObject $reportjson -Encoding utf8
$sw.Stop()
$elapsedsecs = [math]::Round($sw.Elapsed.TotalSeconds)

if (Test-Path .\report.json){
    Write-Host "Report generated successfully. (took $elapsedsecs secs)"
} else {
    Write-Error "Error creating report file."
}

