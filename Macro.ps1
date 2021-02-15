function Set-Shortcut {
param ([string]$SourceLnk, [string]$DestinationPath, [string]$Arguments, [string]$StartIn)
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($SourceLnk)
    $Shortcut.TargetPath = $DestinationPath
	$Shortcut.Arguments = $Arguments
	$Shortcut.WorkingDirectory = $StartIn
    $Shortcut.Save()
}

$CPUArch = Write-Output "$env:PROCESSOR_ARCHITECTURE"
$DownloadLoc = "$env:USERPROFILE\AppData\Local\DCode"
$MacroSource = 'https://github.com/Refr3sh/YWTTC/raw/main/CtinMacro.zip'
$URLPic = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Pic/1.png"
$MacroDest = "$Env:APPDATA\Microsoft\AddIns"


##Background
Start-BitsTransfer -Source "$URLPic" -Destination "$DownloadLoc"

##Macro
Start-BitsTransfer -Source "$MacroSource" -Destination "$DownloadLoc"
Expand-Archive -Path "$DownloadLoc\CtinMacro.zip" -Destination "$MacroDest\"
Remove-Item "$DownloadLoc\CtinMacro.zip" -Force -ErrorAction SilentlyContinue
 Write-Output 'Activate CtinMacro AddIn in Excel'
 Write-Output 'For Using macros, save the Move Task query as .txt in default location (Documents)'
 Write-Output 'Then open EXCEL and press ALT+F1 for barcodes and ALT+F4 For Elvis'
Pause
$LookForSh = Test-Path "$env:USERPROFILE\Desktop\EXCEL.lnk"
If(!($LookForSh -eq $true)){
    If($CPUArch -eq 'AMD64'){
    Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    }else{
    Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    }
}
exit
