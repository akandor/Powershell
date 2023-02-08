$tenantID = "" #705088ec-e683-4174-b9bc-1920644dd49a
$localfile = ""
$sheet = ""

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
}

$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open($localfile)

function Get-Sheets {
    [array]$currentExcelWorkSheets = $ExcelWorkBook.Sheets

    return $currentExcelWorkSheets
}

Clear-Host

Write-Host "Please choose a Workbook Sheet" -ForegroundColor Green
$i=1
$sheetList = Get-Sheets
ForEach($sheetEntry in $sheetList) {
If($sheetEntry.Name -eq $null -Or $sheetEntry.Name -eq "Data") {
    continue
    }
Write-Host("["+ $sheetEntry.Index + "] " + $sheetEntry.Name)
}
[uint16]$sheetID = Read-Host -Prompt "Choose"

$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item($sheetID)

$ii = 0

for($i=2;$i -le $ExcelWorkSheet.UsedRange.Rows.Count;$i++) {
    $user = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(1).Text)
    $number = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(2).Text)
    $cp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(3).Text)
    $dp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(4).Text)
    $vrp = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(5).Text)
    $moh = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(6).Text)
    $vm_lang = ($ExcelWorkSheet.Rows.Item($i).Columns.Item(7).Text)

    $progress = 100 / ($ExcelWorkSheet.UsedRange.Rows.Count - 1) * ($ii++)
    $progressRounded = [math]::Round($progress)

    Write-Progress -Activity "Updating User $user" -Status "$progressRounded% Complete" -PercentComplete $progress

    try {
        Set-CsPhoneNumberAssignment -Identity $user -PhoneNumber $number -PhoneNumberType DirectRouting
        Grant-CsTeamsCallingPolicy -Identity $user -PolicyName $cp
        Grant-CsTenantDialPlan -Identity $user -PolicyName $dp
        Grant-CsOnlineVoiceRoutingPolicy -Identity $user -PolicyName $vrp
        Grant-CsTeamsCallHoldPolicy -Identity $user -PolicyName $moh
        Set-CsOnlineVoicemailUserSettings -Identity $user -PromptLanguage $vm_lang
        Grant-CsTeamsFeedbackPolicy -Identity $user -PolicyName "Disable Survey Policy"
    }
    catch {
        write-host "Phone Number $number cannot be added to $user" -ForegroundColor Yellow 
        $error[0].Exception
        continue
    }

}

$ExcelWorkBook.close()