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
$URLAllFonts = "https://github.com/Refr3sh/YWTTC/raw/main/Font.zip"
$URLFont = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Font.ps1"
$URLQuiet = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/Quiet.ps1"
$MacroSource = 'https://github.com/Refr3sh/YWTTC/raw/main/CtinMacro.zip'
$MacroDest = "$Env:APPDATA\Microsoft\AddIns"

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
    $web = New-Object System.Net.WebClient
	$web.DownloadString("$URLQuiet") > "$DownloadLoc\Quiet.ps1"
    Remove-Variable web
}

##Font
$ChkFont = Test-Path "$DownloadLoc\Font.ps1"
If(!($ChkFont -eq $True)){
    $web = New-Object System.Net.WebClient
	$web.DownloadString("$URLFont") > "$DownloadLoc\Font.ps1"
    Start-BitsTransfer -Source $URLAllFonts -Destination $DownloadLoc
	Expand-Archive -LiteralPath "$DownloadLoc\Font.zip" -DestinationPath "$DownloadLoc\Font"
	Remove-Item "$DownloadLoc\Font.zip" -Force
    Remove-Variable web
}

##ExcelMacro
$LookForMacro = Test-Path "$MacroDest\CtinMacro.xlam"
#$LookForOM = Test-Path "$MacroDest\OMTool*.xlam"
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
    If($CPUArch -eq 'AMD64'){
    Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
    }else{
    Set-Shortcut "$Env:USERPROFILE\Desktop\EXCEL.lnk" "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    }
}
Start-Sleep 5
&"$DownloadLoc\Font.ps1" -WindowStyle Hidden
Start-Sleep 5
&"$DownloadLoc\Quiet.ps1" -WindowStyle Hidden
