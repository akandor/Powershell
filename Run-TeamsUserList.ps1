$tenantID = ""
$localfile = ""

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

function Get-Sheets {
    [array]$currentExcelWorkSheets = $ExcelWorkBook.Sheets

    return $currentExcelWorkSheets
}

Clear-Host

while($sheetID -eq 0) {
    Clear-Host
    Write-Host "Please choose a Workbook Sheet" -ForegroundColor Green
    $i=1
    foreach ($sheet in $e.workbook.worksheets) {
        If($null -eq $sheet.name -Or $sheet.name -eq "Data" -Or $sheet.name -eq "Template" -Or $sheet.name -eq "Instructions") {
            continue
            }
        Write-Host("["+ $i++ + "] " + $sheet.name)
    }
    $sheetID = Read-Host -Prompt "Choose"
}

$WorkSheets = $e.workbook.worksheets[$sheetID].Cells

[array]$userData = Import-Excel -Path $localfile -WorksheetName $e.workbook.worksheets[$sheetID].Name

Clear-Host

$i = 1

foreach($userEntry in $userData) {

    $progress = 100 / ($userData.Count) * ($i)
    $progressRounded = [math]::Round($progress)

    $user = $userEntry.UPN

    $user = $userEntry.UPN
    $number = $userEntry.Number
    $cp = $userEntry.'Calling Policy'
    $dp = $userEntry.'Dial Plan'
    $vrp = $userEntry.'Voice Routing Policy'
    $moh = $userEntry.'Call Hold Policy'
    $cpp = $userEntry.'Call Park Policy'
    $cip = $userEntry.'Caller ID Policy'
    $emp = $userEntry.'Emergency Policiy'
    $vm_lang = $userEntry.'Voice Mail'

    Write-Progress -Activity "Updating User $user" -Id 1 -Status "$progressRounded% Complete" -PercentComplete $progress

    $i++

    if($userEntry.Generated -eq "y") {
        continue
    }

    try {
        if($user -ne "" -And $null -ne $user) {
            if($number -ne "" -And $number -ne $null) {
                Set-CsPhoneNumberAssignment -Identity $user -PhoneNumber $number -PhoneNumberType DirectRouting -ErrorActio Stop
            }
            if($cp -ne "" -And $cp -ne $null) {
                Grant-CsTeamsCallingPolicy -Identity $user -PolicyName $cp -ErrorActio Stop
            }
            if($dp -ne "" -And $dp -ne $null) {
                Grant-CsTenantDialPlan -Identity $user -PolicyName $dp -ErrorActio Stop
            }
            if($vrp -ne "" -And $vrp -ne $null) {
                Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName $vrp -ErrorActio Stop
            }
            if($moh -ne "" -And $moh -ne $null) {
                Grant-CsTeamsCallHoldPolicy -Identity $user -PolicyName $moh -ErrorActio Stop
            }
            if($vm_lang -ne "" -And $vm_lang -ne $null) {
                Set-CsOnlineVoicemailUserSettings -Identity $user -PromptLanguage $vm_lang -ErrorActio Stop
            }
            if($cpp -ne "" -And $cpp -ne $null) {
                Grant-CsTeamsCallParkPolicy -Identity $user -PolicyName $cpp -ErrorActio Stop
            }
            if($cip -ne "" -And $cip -ne $null) {
                Grant-CsCallingLineIdentity -Identity $user -PolicyName $cip -ErrorActio Stop
            }
            if($emp -ne "" -And $emp -ne $null) {
                Grant-CsTeamsEmergencyCallingPolicy -Identity $user -PolicyName $emp -ErrorActio Stop
            }
            Grant-CsTeamsFeedbackPolicy -Identity $user -PolicyName "Disable Survey Policy" -ErrorActio Stop

            $WorkSheets[($i),11].Value = "y"
        }
        
    }
    catch {
        write-host "Phone Number $number cannot add to $user" -ForegroundColor Yellow 
        $error[0].Exception
        continue
    }

}

Write-Progress -Activity "Updating User $user" -Id 1 -Completed

Close-ExcelPackage $e

$message ="All done! Press any key to quit."
pause ($message)