# My Powershell Library

## Run-/Create-TeamsUserList.ps1
Use the Tempalte Excel file to set policies for enterprise voice enabled users.

### Steps
1. Run the Create-TeamsUserList.ps1 and select the Template File. The script will update the data sheet with the configured Policies in MS Teams. Optionally you can download all enterprise voice enabled users from the tenant and create a sheet.
2. After you set up the policies for every user, you can run the Run-TeamsUserList.ps1 script and select the file and the sheet you want to use.

## PopulateTeamsHolidays.ps1
Update/Create the Teams holiday schedule with new events.

### Steps
Run the PopulateTeamsHoliday.ps1 script. Decide by yourself to use a CSV local file or use the online API from https://www.openholidaysapi.org. This will then create a new Holiday table from today plus 1 year in MS Teams Voice.

## AudioCodesSBCPublicNATIP.ps1
Update Audiocodes SBC NAT Public IP for non static public ip

## TeamsModule.ps1
Create Incremental INI File for Audiocodes SBC and MS Teams Direct Routing
