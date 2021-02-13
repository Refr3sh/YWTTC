$Itm = "$env:TEMP", "$env:UserProfile\Appdata\Roaming\Microsoft\Windows\Recent\", "$env:UserProfile\Appdata\Roaming\Microsoft\Windows\Cookies\", "$env:UserProfile\AppData\Roaming\Microsoft\Office\Recent\"
    Get-ChildItem  $Itm -Recurse -ErrorAction SilentlyContinue |  foreach {Remove-Item $_.fullname -Force -Recurse -ErrorAction SilentlyContinue}
Start-Sleep 5
$makequiet = "*authman*", "*concen*", "*wfcrun*", "*redirect*", "*selfservicePlug*", "*receiver*", "*chrome*", "*lync*", "*teams*", "*pnamain*"
  foreach( $process in $makequiet )
   {
    Stop-Process -name $makequiet -Force -Verbose
   }