$tenantID = ""

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

Write-Host "Do you want remove or add a number" -ForegroundColor Green
Write-Host "[1] Add/Update a number (Default)" -ForegroundColor Green
Write-Host "[2] Remove a number" -ForegroundColor Green
$updateRemove = Read-Host -Prompt "Choose"

Clear-Host

Write-Host "Please enter the UPN/Identity of the user." -ForegroundColor Green
$upn = Read-Host -Prompt "UPN/Identiy"

Clear-Host

Write-Host "Please enter the phone number (E.164 Format)." -ForegroundColor Green
$upn = Read-Host -Prompt "Phone Number"

if($importUsers -eq "2") {
    Clear-Host
}