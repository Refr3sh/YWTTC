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
If($CPUArch -eq 'AMD64'){
    Set-Shortcut "$StartFolder\Quiet.lnk" 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File .\DCode\Quiet.ps1' '%localappdata%'
}else{
    Set-Shortcut "$StartFolder\Quiet.lnk" 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File .\DCode\Quiet.ps1' '%localappdata%'
}
$Itm = "$env:TEMP", "$env:UserProfile\Appdata\Roaming\Microsoft\Windows\Recent\", "$env:UserProfile\Appdata\Roaming\Microsoft\Windows\Cookies\", "$env:UserProfile\AppData\Roaming\Microsoft\Office\Recent\"
    Get-ChildItem  $Itm -Recurse -ErrorAction SilentlyContinue |  foreach {Remove-Item $_.fullname -Force -Recurse -ErrorAction SilentlyContinue}
$makequiet = "*authman*", "*concen*", "*wfcrun*", "*redirect*", "*selfservicePlug*", "*receiver*", "*chrome*", "*lync*", "*teams*", "*pnamain*"
  foreach( $process in $makequiet )
   {
    Stop-Process -name $makequiet -Force -ErrorAction SilentlyContinue
   }
exit
