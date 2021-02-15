$web = New-Object System.Net.WebClient
$DownloadLoc = "$env:USERPROFILE\AppData\Local\DCode"
$MacroDest = "$Env:APPDATA\Microsoft\AddIns"
$EXCEL = Get-Process EXCEL -ErrorAction SilentlyContinue
If($EXCEL){
  $EXCEL.CloseMainWindow()
  Sleep 10
  If(!$EXCEL.HasExited){
    $EXCEL | Stop-Process -Force
  }
}
Remove-Variable EXCEL


Remove-Item "$MacroDest\CtinMacro.*" -Force -ErrorAction SilentlyContinue
Remove-Item "$DownloadLoc\CtinMacro.*" -Force -ErrorAction SilentlyContinue
Remove-Item "$DownloadLoc\*.ps1" -Force -ErrorAction SilentlyContinue
##DownloadFolder
$ChkDownloadLoc = Test-Path "$DownloadLoc"
If(!($ChkDownloadLoc -eq $True)){
	New-Item -Path "$DownloadLoc" -ItemType Directory -Force
	New-Item -Path "$DownloadLoc\Font" -ItemType Directory -Force
}
$web.DownloadString('https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Main.ps1') > "$DownloadLoc\Main.ps1"
&"$Env:UserProfile\Appdata\Local\DCode\Main.ps1" -WindowStyle Hidden
Remove-Variable web
exit
