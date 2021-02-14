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
$StartFolder = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
$DownloadLoc = "$env:USERPROFILE\AppData\Local\DCode"
$web = New-Object System.Net.WebClient
$URLAllFonts = "https://github.com/Refr3sh/YWTTC/raw/main/Font.zip"
$URLFont = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Font.ps1"
$URLQuiet = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Quiet.ps1"
$MacroSource = 'https://github.com/Refr3sh/YWTTC/raw/main/CtinMacro.zip'
$MacroDest = "$Env:APPDATA\Microsoft\AddIns"



##SetShortcuts
If($CPUArch -eq 'AMD64'){
	$ShortcutQuiet = Set-Shortcut "$StartFolder\Quiet.lnk" 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File ".\DCode\Quiet.ps1"' '%localappdata%'
	$ShortcutFont = Set-Shortcut "$StartFolder\Font.lnk" 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File ".\DCode\Font.ps1"' '%localappdata%'
	$EXCELShortcut = Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
}else{
	$ShortcutQuiet = Set-Shortcut "$StartFolder\Quiet.lnk" 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File ".\DCode\Quiet.ps1"' '%localappdata%'
	$ShortcutFont = Set-Shortcut "$StartFolder\Font.lnk" 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File ".\DCode\Font.ps1"' '%localappdata%'
	$EXCELShortcut = Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
}

##ClearStart
Remove-Item "$StartFolder\*" -Force

##DownloadFolder
$ChkDownloadLoc = Test-Path $DownloadLoc
If(!($ChkDownloadLoc -eq $True)){
	New-Item -Path $DownloadLoc -ItemType Directory -Force
	New-Item -Path "$DownloadLoc\Font" -ItemType Directory -Force
}

##Quiet
$ChkQuiet = Test-Path "$DownloadLoc\Quiet.ps1"
If(!($ChkQuiet -eq $True)){
	$web.DownloadString("$URLQuiet") > "$DownloadLoc\Quiet.ps1"
}
If($ChkQuiet -eq $True){
#	&"$DownloadLoc\Quiet.ps1" -WindowStyle Hidden
	$ShortcutQuiet
}else{
	Write-Output 'Something went wrong while cleaning up'
	Pause
	Exit
}

##Font
$ChkFont = Test-Path "$DownloadLoc\Font.ps1"
If(!($ChkFont -eq $True)){
	$web.DownloadString("$URLFont") > "$DownloadLoc\Font.ps1"
	Start-BitsTransfer -Source $URLAllFonts -Destination $DownloadLoc
	Expand-Archive -LiteralPath "$DownloadLoc\Font.zip" -DestinationPath "$DownloadLoc\Font"
	Remove-Item "$DownloadLoc\Font.zip" -Force
}
If($ChkFont -eq $True){
	&"$DownloadLoc\Font.ps1" -WindowStyle Hidden
	$ShortcutFont
}else{
	Write-Output 'Something went wrong while installing fonts'
	Pause
	Exit
}

##ExcelMacro
$LookForMacro = Test-Path "$MacroDest\CtinMacro.xlam"
$LookForOM = Test-Path "$MacroDest\OMTool*.xlam"
$LookForSh = Test-Path "$env:USERPROFILE\Desktop\EXCEL.lnk"
If(!($LookForMacro -eq $true)){
	Start-BitsTransfer -Source $MacroSource -Destination $MacroDest
	Expand-Archive -LiteralPath "$MacroDest\CtinMacro.zip" -DestinationPath $MacroDest
	Remove-Item "$MacroDest\CtinMacro.zip" -Force
	Write-Output 'Activate CtinMacro AddIn in Excel'
    Write-Output 'For Using macros, save the Move Task query as .txt in default location (Documents)'
    Write-Output 'Then open EXCEL and press CTRL+F1 for barcodes and CTRL+F2 For Elvis'
    Pause
}
<#
If(!($LookForOM -eq $true)){
	Robocopy "\\gbhgmercser0021\pdg\" "$MacroDest" "OMTool*.xlam" /FFT /Z /W:5
}
#>
If(!($LookForSh -eq $true)){
	$EXCELShortcut
}
