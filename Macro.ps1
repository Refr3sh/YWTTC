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
$MacroSource = 'https://github.com/Refr3sh/YWTTC/raw/main/CtinMacro.zip'
$MacroDest = "$Env:APPDATA\Microsoft\AddIns"
$LookForMacro = Test-Path "$MacroDest\CtinMacro.xlam"
$LookForSh = Test-Path "$env:USERPROFILE\Desktop\EXCEL.lnk"

If(!($LookForMacro -eq $true)){
	Start-BitsTransfer -Source $MacroSource -Destination $MacroDest
	Expand-Archive -Path "$MacroDest\CtinMacro.zip" -Destination "$MacroDest\"
	Remove-Item "$MacroDest\CtinMacro.zip" -Force -ErrorAction SilentlyContinue
	Write-Output 'Activate CtinMacro AddIn in Excel'
    Write-Output 'For Using macros, save the Move Task query as .txt in default location (Documents)'
    Write-Output 'Then open EXCEL and press CTRL+F1 for barcodes and CTRL+F2 For Elvis'
    Pause
}

If(!($LookForSh -eq $true)){
    If($CPUArch -eq 'AMD64'){
    Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    }else{
    Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    }
}
exit
