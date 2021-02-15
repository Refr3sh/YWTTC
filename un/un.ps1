function Set-Shortcut {
param ([string]$SourceLnk, [string]$DestinationPath, [string]$Arguments, [string]$StartIn)
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($SourceLnk)
    $Shortcut.TargetPath = $DestinationPath
	$Shortcut.Arguments = $Arguments
	$Shortcut.WorkingDirectory = $StartIn
    $Shortcut.Save()
}

$URLFontUninst = "https://raw.githubusercontent.com/Refr3sh/YWTTC/main/un/Uninst.ps1"
$StartFolder = "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
$DownloadLoc = "$env:USERPROFILE\AppData\Local\DCode"
$MacroDest = "$Env:APPDATA\Microsoft\AddIns"
$web = New-Object System.Net.WebClient

Remove-Item "$StartFolder\Font.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item "$StartFolder\Quiet.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item "$DownloadLoc\*" -Force -Recurse -ErrorAction SilentlyContinue
$web.DownloadString("$URLFontUninst") > "$env:USERPROFILE\AppData\Local\Uninst.ps1"

If($CPUArch -eq 'AMD64'){
    Set-Shortcut "$StartFolder\u.lnk" 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninst.ps1' '%localappdata%'
}else{
    Set-Shortcut "$StartFolder\u.lnk" 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninst.ps1' '%localappdata%'
}

$ProcessName = "EXCEL"
if((get-process $ProcessName -ErrorAction SilentlyContinue) -eq $Null){
	Remove-Item "$MacroDest\CtinMacro.xlam" -Force -ErrorAction SilentlyContinue
}else{
	Write-Output "Please close Excel"
	break
	Remove-Item "$MacroDest\CtinMacro.xlam" -Force -ErrorAction SilentlyContinue
}
exit
