$tenantID = ""
$localFile = ""

$vrp = ""
$dp = ""
$cp = ""
$moh = ""
$vm_lang = "de-DE"

function Get-VoiceRoutingPolicies {
    [array]$getAllVoiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy

    return $getAllVoiceRoutingPolicies
}

function Get-DialPlans {
    [array]$getAllDialPlans = Get-CsTenantDialPlan

    return $getAllDialPlans
}

function Get-CallingPolicies {
    [array]$getAllCallingPolicies = Get-CsTeamsCallingPolicy

    return $getAllCallingPolicies
}

function Get-CallHoldPolicies {
    [array]$getAllCallHoldPolicies = Get-CsTeamsCallHoldPolicy

    return $getAllCallHoldPolicies
}

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

Clear-Host

if($vrp -eq "") {
    Write-Host "Please choose a Voice Routing Policy" -ForegroundColor Green
    $i=1
    $vrpList = Get-VoiceRoutingPolicies
    ForEach($vrpEntry in $vrpList) {
    If($vrpEntry.Identity -eq $null) {
        continue
        }
    Write-Host("["+ $i++ + "] " + $vrpEntry.Identity -replace 'Tag:','')
    }
    $vrpID = Read-Host -Prompt "Choose"
    $vrp = $vrpList[$vrpID-1].Identity -replace 'Tag:',''
    Write-Host($vrp)
}

Clear-Host

if($dp -eq "") {
    Write-Host "Please choose a Dial Plan" -ForegroundColor Green
    $i=1
    $dpList = Get-DialPlans
    ForEach($dpEntry in $dpList) {
    If($dpEntry.Identity -eq $null) {
        continue
        }
    Write-Host("["+ $i++ + "] " + $dpEntry.Identity -replace 'Tag:','')
    }
    $dpID = Read-Host -Prompt "Choose"
    $dp = $dpList[$dpID - 1].Identity -replace 'Tag:',''
    Write-Host($dp)
}

Clear-Host

if($cp -eq "") {
    Write-Host "Please choose a Calling Policy" -ForegroundColor Green
    $i=1
    $cpList = Get-CallingPolicies
    ForEach($cpEntry in $cpList) {
    If($cpEntry.Identity -eq $null) {
        continue
        }
    Write-Host("["+ $i++ + "] " + $cpEntry.Identity -replace 'Tag:','')
    }
    $cpID = Read-Host -Prompt "Choose"
    $cp = $cpList[$cpID - 1].Identity -replace 'Tag:',''
    Write-Host($cp)
}

Clear-Host

if($moh -eq "") {
    Write-Host "Please choose a Call Hold Policy" -ForegroundColor Green
    $i=1
    $mohList = Get-CallHoldPolicies
    ForEach($mohEntry in $mohList) {
    If($mohEntry.Identity -eq $null) {
        continue
        }
    Write-Host("["+ $i++ + "] " + $mohEntry.Identity -replace 'Tag:','')
    }
    $mohID = Read-Host -Prompt "Choose"
    $moh = $mohList[$mohID - 1].Identity -replace 'Tag:',''
    Write-Host($moh)
}

break

#$user = "xDamovoTeams1@staedtler.com"
$path = "G:\Region-Sued-Sales\s\staedtler\nürnberg\aufträge\221020 P20004920 V20009886 sf-21876 ms-stemas als uc-platform\PM_Stefan Bahnsen\Projektdokumente Projektteam\DE\Userliste_Nuernberg_07_02_de2.csv"

$ErrorActionPreference = "Stop"

try {
    $csv = import-csv $path -Delimiter ","
}
catch {
    write-host "CSV kann nicht importiert werden" -ForegroundColor Yellow
    $error[0].Exception
    break
}




foreach ($zeile in $csv) {
    $user = $zeile.upn
    $nummer= $zeile.Rufnummer
    try {
        Set-CsPhoneNumberAssignment -Identity $user -PhoneNumber $nummer -PhoneNumberType DirectRouting
        Grant-CsTeamsCallingPolicy -Identity $user -PolicyName $cp
        Grant-CsTenantDialPlan -Identity $user -PolicyName $dp
        Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName $vrp
        Grant-CsTeamsCallHoldPolicy -Identity $user -PolicyName $moh
        Set-CsOnlineVoicemailUserSettings -Identity $user -PromptLanguage $vm_lang
        Grant-CsTeamsFeedbackPolicy -Identity $user -PolicyName "Disable Survey Policy"
    }
    catch {
        write-host "Rufnummer $nummer kann nicht dem User $user zugewiesen werden" -ForegroundColor Yellow 
        $error[0].Exception
        continue
    }
}

get-csonlineuser -Filter 'EnterpriseVoiceEnabled -eq $true' | select userprincipalname, lineuri, TenantDialplan, OnlineVoiceRoutingPolicy, TeamsCallingPolicy | ft

$user = "ssd61209@staedtler.onmicrosoft.com"
get-csonlineuser -Identity $user | select userprincipalname, lineuri, TenantDialplan, OnlineVoiceRoutingPolicy, TeamsCallingPolicy, TeamsCallHoldPolicy, TeamsIPPhonePolicy| ft


Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName vrp-de-international