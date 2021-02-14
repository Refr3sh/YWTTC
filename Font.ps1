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
    Set-Shortcut "$StartFolder\Font.lnk" 'C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File .\DCode\Font.ps1' '%localappdata%'
}else{
    Set-Shortcut "$StartFolder\Font.lnk" 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe' '-ExecutionPolicy Bypass -WindowStyle Hidden -File .\DCode\Font.ps1' '%localappdata%'
}
Add-Type -Name Session -Namespace "" -Member @"
[DllImport("gdi32.dll")]
public static extern int AddFontResource(string filePath);
"@

$null = foreach($font in Get-ChildItem "$env:USERPROFILE\AppData\Local\DCode\Font\"-Recurse -Include *.ttf, *.otf) {
    [Session]::AddFontResource($font.FullName)
}
exit
