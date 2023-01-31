 $LF = "`r`n";

# AudioCodes SBC Data
$ip = "[IP-Address]"
$username = "[Username]"
$password = "[Password]"

# NAT Settings
$srcPortStart = "7000"
$srcPortEnd = "7499"

# Get Public IP
$mypublicip = (Invoke-WebRequest -uri "http://ifconfig.me/ip").Content

# Build CLI Incremental String
$cliData =  "configure network$LF nat-translation 0$LF src-interface-name `"IF-WAN`"$LF target-ip-address `"$mypublicip`"$LF src-start-port `"$srcPortStart`"$LF src-end-port `"$srcPortEnd`"$LF activate"

# CLI URL
$URLFull = "http://{0}/api/v1/files/cliScript" ` -f $ip
$URLIncremental = "http://{0}/api/v1/files/cliScript/incremental" ` -f $ip

# REST API Authentication
$authHash = [Convert]::ToBase64String( ` [Text.Encoding]::ASCII.GetBytes( ` ("{0}:{1}" -f $username,$password)))

# INI File Body
$boundary = [System.Guid]::NewGuid().ToString(); 
$bodyLines = (
    "--$boundary",
("Content-Disposition: form-data; name=`"file`";" + `
" filename=`"file.txt`""),
"Content-Type: application/octet-stream$LF", $cliData,
"--$boundary--$LF"
) -join $LF

# Get Full INI File
$response = Invoke-RestMethod -Uri $URLFull -Method Get ` -Headers @{Authorization=("Basic {0}" -f $authHash)} ` 
#$response

# Get target-ip-address
$targetIPString = ($response  |  Select-String -Pattern "target-ip-address `"\d{1,3}(\.\d{1,3}){3}`"" -AllMatches).Matches.Value

# Check and Set New Public IP
if($targetIPString) {
    $targetIP = ($targetIPString  |  Select-String -Pattern "\d{1,3}(\.\d{1,3}){3}" -AllMatches).Matches.Value
    if($targetIPString -ne $mypublicip) {
        $changeIP = Invoke-RestMethod -Uri $URLIncremental -Method Put ` -Headers @{Authorization=("Basic {0}" -f $authHash)} ` -ContentType "multipart/form-data; boundary=$boundary" ` -Body $bodyLines
        $changeIP | ConvertTo-Json
    }
} 