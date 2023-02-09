$tenantID = "705088ec-e683-4174-b9bc-1920644dd49a" #705088ec-e683-4174-b9bc-1920644dd49a
$localfile = ""
$sheet = ""
[uint16]$sheetID = 0
$global:progress = 0

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
    $ExcelObj = New-Object -comobject Excel.Application
} catch {
    Write-Host "Error while opening Excel. Please be sure that Excel is already installed and restart the script!" -ForegroundColor Red
    $error[0].Exception
    Break
}

try {
    $ExcelWorkBook = $ExcelObj.Workbooks.Open($localfile)
} catch {
    Write-Host "Error while opening file. Please be sure that Excel is already installed and restart the script!" -ForegroundColor Red
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
    $sheetList = Get-Sheets
    ForEach($sheetEntry in $sheetList) {
    If($null -eq $sheetEntry.Name -Or $sheetEntry.Name -eq "Data") {
        continue
        }
    Write-Host("["+ $sheetEntry.Index + "] " + $sheetEntry.Name)
    }
    $sheetID = Read-Host -Prompt "Choose"
}

$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($sheetID)

Clear-Host

$ii = 1

for($i=2;$i -le $ExcelWorkSheet.UsedRange.Rows.Count;$i++) {

    $global:progress = 100 / ($ExcelWorkSheet.UsedRange.Rows.Count - 1) * ($ii++)
    $progressRounded = [math]::Round($global:progress)

    $user = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(1).Text)
    $number = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(2).Text)
    $cp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(3).Text)
    $dp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(4).Text)
    $vrp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(5).Text)
    $moh = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(6).Text)
    $cpp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(7).Text)
    $cip = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(8).Text)
    $emp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(9).Text)
    $vm_lang = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(10).Text)

    Write-Progress -Activity "Updating User $user" -Status "$progressRounded% Complete" -PercentComplete $global:progress

    if($ExcelWorkSheet.Rows.Item($i).Columns.Item(11).Text -eq "y") {
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

            $ExcelWorkSheet.Rows.Item($i).Columns.Item(11) = "y"
        }
        
    }
    catch {
        write-host "Phone Number $number cannot add to $user" -ForegroundColor Yellow 
        $error[0].Exception
        continue
    }

}

$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)  