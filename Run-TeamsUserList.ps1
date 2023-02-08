Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') Filter = 'SpreadSheet (*.xlsx)|*.xlsx' }
$null = $FileBrowser.ShowDialog()

$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("TeamsUserList.xlsx")

[array]$currentExcelWorkSheets = $ExcelWorkBook.Sheets