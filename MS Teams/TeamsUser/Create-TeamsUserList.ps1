$tenantID = ""
$localfile = ""
$sheetName = ""

Set-StrictMode -Version "2.0"

$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

try {
Import-Module -Name ($ScriptDirectory + "\lib\ImportExcel")
}
catch {
Write-Host "Error while loading supporting PowerShell Scripts"
}

$sheet = ""
[uint16]$sheetID = 0

Clear-Host

if($tenantID -eq "") {
    Write-Host "Please enter your Tenant ID (can be left empty if you have only one)" -ForegroundColor Green
    $tenantID= Read-Host -Prompt "Tenant ID"
    If($tenantID.Length -gt 10) {
        Connect-MicrosoftTeams -TenantId $tenantID
    } else {
        Connect-MicrosoftTeams
    }
} else {
    Connect-MicrosoftTeams -TenantId $tenantID
}

if($localfile -eq "") {
    Add-Type -AssemblyName System.Windows.Forms

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
    }
    $null = $FileBrowser.ShowDialog()
    $localfile = $FileBrowser.FileName

    if($localfile -eq "" -Or $localfile -eq $null) {
        Write-Host "You need to select a file. Please restart the script!" -ForegroundColor Red
        Break
    }
}

try {
    $e = Open-ExcelPackage $localfile
} catch {
    Write-Host "Error while opening Excel file!" -ForegroundColor Red
    $error[0].Exception
    Break
}

Write-Host "Do you want to import all Enterprise Voice enabled users?" -ForegroundColor Green
$importUsers = Read-Host -Prompt "Y/N"

if($sheetName -eq "" -And ($importUsers -eq "Y" -Or $importUsers -eq "y")) {
    Clear-Host
    
    Write-Host "Please enter a sheet name for the import." -ForegroundColor Green
    $sheetName = Read-Host -Prompt "Sheet Name"

    Add-Worksheet -ExcelPackage $e -WorksheetName $sheetName -CopySource $e.workbook.worksheets["Template"] -MoveToStart
}

$DataWorkSheet = $e.workbook.worksheets["Data"].Cells

function Get-AllEnterpriseVoiceUsers {
    [array]$enterpriseVoiceUsers = Get-CSOnlineUser -Filter 'EnterpriseVoiceEnabled -eq $true' | select userprincipalname, lineuri, TenantDialplan, OnlineVoiceRoutingPolicy, TeamsCallingPolicy, TeamsCallHoldPolicy, TeamsCallParkPolicy, CallingLineIdentity, TeamsEmergencyCallingPolicy
    return $enterpriseVoiceUsers
}
function Get-CallingPolicies {
    [array]$callingPolicies = Get-CsTeamsCallingPolicy
    return $callingPolicies
}
function Get-DialPlans {
    [array]$dialPlans = Get-CsTenantDialPlan
    return $dialPlans
}

function Get-VoiceRoutingPolicies {
    [array]$voiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy
    return $voiceRoutingPolicies
}
function Get-CallHoldPolicies {
    [array]$callHoldPolicies = Get-CsTeamsCallHoldPolicy
    return $callHoldPolicies
}
function Get-CallParkPolicies {
    [array]$callParkPolicies = Get-CsTeamsCallParkPolicy
    return $callParkPolicies
}
function Get-CallingLineIdentities {
    [array]$callingLineIdentities = Get-CsCallingLineIdentity
    return $callingLineIdentities
}
function Get-EmergencyCallingPolicies {
    [array]$emergencyCallingPolicies = Get-CsTeamsEmergencyCallingPolicy
    return $emergencyCallingPolicies
}

if($importUsers -eq "Y" -Or $importUsers -eq "y") {
    $UserWorkSheet = $e.workbook.worksheets[$sheetName].Cells
    $i = 2

    foreach($entry in Get-AllEnterpriseVoiceUsers) {
        Write-Host($entry.UserPrincipalName)

        $e.workbook.worksheets[$sheetName].InsertRow($i,1)

        $UserWorkSheet[($i),1].Value = $entry.UserPrincipalName -replace "Tag:",""
        $UserWorkSheet[($i),2].Value = $entry.LineURI -replace "tel:",""
        $UserWorkSheet[($i),3].Value = $entry.TeamsCallingPolicy
        $UserWorkSheet[($i),4].Value = $entry.TenantDialplan
        $UserWorkSheet[($i),5].Value = $entry.OnlineVoiceRoutingPolicy
        $UserWorkSheet[($i),6].Value = $entry.TeamsCallHoldPolicy
        $UserWorkSheet[($i),7].Value = $entry.TeamsCallParkPolicy
        $UserWorkSheet[($i),8].Value = $entry.CallingLineIdentity
        $UserWorkSheet[($i),9].Value = $entry.TeamsEmergencyCallingPolicy
        $i++
    }
}


$i = 2

foreach($entry in Get-CallingPolicies) {
    $DataWorkSheet[($i),1].Value = $entry.Identity -replace "Tag:",""
    $i++
}

$i = 2

foreach($entry in Get-DialPlans) {
    $DataWorkSheet[($i),2].Value = $entry.Identity -replace "Tag:",""
    $i++
}

$i = 2

foreach($entry in Get-VoiceRoutingPolicies) {
    $DataWorkSheet[($i),3].Value = $entry.Identity -replace "Tag:",""
    $i++
}

$i = 2

foreach($entry in Get-CallHoldPolicies) {
    $DataWorkSheet[($i),4].Value = $entry.Identity -replace "Tag:",""
    $i++
}

$i = 2

foreach($entry in Get-CallParkPolicies) {
    $DataWorkSheet[($i),5].Value = $entry.Identity -replace "Tag:",""
    $i++
}

$i = 2

foreach($entry in Get-CallingLineIdentities) {
    $DataWorkSheet[($i),6].Value = $entry.Identity -replace "Tag:",""
    $i++
}

$i = 2

foreach($entry in Get-EmergencyCallingPolicies) {
    $DataWorkSheet[($i),7].Value = $entry.Identity -replace "Tag:",""
    $i++
}

Close-ExcelPackage $e