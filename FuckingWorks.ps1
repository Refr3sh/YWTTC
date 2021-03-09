Function Write-HostCenter { param($Message) Write-Host ("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($Message.Length / 2)))), $Message)}
Function Write-HostCenterR { param($Message) Write-Host ("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($Message.Length / 2)))), $Message) -ForegroundColor Red}
Function Write-HostCenterG { param($Message) Write-Host ("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($Message.Length / 2)))), $Message) -ForegroundColor Green}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public class Win32SetWindow {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
"@
$username="ndcdespatch"
$password="Ndcdespatch!2"
$ie = new-object -comobject InternetExplorer.Application;
$ie.visible = $true;
$ie.MenuBar = $False;
$ie.ToolBar = 0;
$ie.top = 15; $ie.width = 630; $ie.height = 200; 
     [Win32SetWindow]::SetForegroundWindow($ie.HWND)

$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=qs');
while($ie.Busy){Start-Sleep -Milliseconds 500}
&"$env:UserProfile\appdata\local\1.lnk" -windowstyle hidden
$login = {
$chk = $ie.Document.IHTMLDocument3_getElementsByTagName("input") | Select-Object Type,Name | Select -Last 1
if($chk.name -eq "DM_TOKEN"){ 
$ie.Document.IHTMLDocument3_getElementById("userId").value = $username
$ie.Document.IHTMLDocument3_getElementById("password").value = $password
$logbutton = $ie.Document.IHTMLDocument3_getElementById("loginButton")
$logbutton.click()
while($ie.Busy){Start-Sleep -Milliseconds 500}
}
}
&$login
$Upallet = ""
while($Upallet -ne "Q")
{
clear
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenter ''
Write-HostCenterG 'Scan pallet or press "Q" to quit'
$Upallet = Read-Host ' '

If($Upallet -like "*q*"){
	$ie.Quit()
	spps -n iexplore -ErrorAction SilentlyContinue
    spps -n ielowutil -ErrorAction SilentlyContinue
    spps -n conhost -ErrorAction SilentlyContinue
    spps -n powershell -ErrorAction SilentlyContinue
    exit}
while(!(Test-Path "$env:UserProfile\1.csv")){
Write-Output 'No .csv file found!'
Start-Sleep 2}

&$login
If($Upallet -like "46*" -or "SYW*"){
If(Test-Path "$env:UserProfile\1.csv"){
$var = Import-Csv -Path "$env:UserProfile\1.csv" | ? {$_.ContainerId -eq $Upallet} | select -First 1
if($var -eq $null){$var = Import-Csv -Path "$env:UserProfile\1.csv" | ? {$_.PalletId -eq $Upallet} | select -First 1}
if($var.OrderId){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=qs&text=' + $var.OrderId + '&textType=0&recentDayCount=14&renderedRetailerId=7001')
while($ie.Busy){Start-Sleep -Milliseconds 500}
$dhlc = $ie.Document.url
$dhlc = $dhlc -replace '\D+([0-9]*).*','$1'
if($dhlc -like "6*"){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=print_labels&showErrors=false&consignmentId=' + $dhlc)
while($ie.Busy){Start-Sleep -Milliseconds 500}
Start-Sleep 2.9
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')}else{
$var = $Upallet
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=qs&text=' + $var + '&textType=10&recentDayCount=14&renderedRetailerId=7001')
while($ie.Busy){Start-Sleep -Milliseconds 500}
$dhlc = $ie.Document.url
$dhlc = $dhlc -replace '\D+([0-9]*).*','$1'
if($dhlc -like "6*"){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=print_labels&showErrors=false&consignmentId=' + $dhlc)
while($ie.Busy){Start-Sleep -Milliseconds 500}
Start-Sleep 2.9
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')}else{Write-HostCenterG 'Nothing Found!!'; sleep 2}}
}else{
$var = $Upallet
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=qs&text=' + $var + '&textType=10&recentDayCount=14&renderedRetailerId=7001')
while($ie.Busy){Start-Sleep -Milliseconds 500}
$dhlc = $ie.Document.url
$dhlc = $dhlc -replace '\D+([0-9]*).*','$1'
if($dhlc -like "6*"){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=print_labels&showErrors=false&consignmentId=' + $dhlc)
while($ie.Busy){Start-Sleep -Milliseconds 500}
Start-Sleep 2.9
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')}else{Write-HostCenterG 'Nothing Found!!'; sleep 2}}
}else{
$var = $Upallet
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=qs&text=' + $var + '&textType=10&recentDayCount=14&renderedRetailerId=7001')
while($ie.Busy){Start-Sleep -Milliseconds 500}
$dhlc = $ie.Document.url
$dhlc = $dhlc -replace '\D+([0-9]*).*','$1'
if($dhlc -like "6*"){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=print_labels&showErrors=false&consignmentId=' + $dhlc)
while($ie.Busy){Start-Sleep -Milliseconds 500}
Start-Sleep 2.9
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')}else{Write-HostCenterG 'Nothing Found!!'; sleep 2}}
}
If($Upallet -like "51*"){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=qs&text=' + $Upallet + '&textType=0&recentDayCount=14&renderedRetailerId=7001')
while($ie.Busy){Start-Sleep -Milliseconds 500}
$dhlc = $ie.Document.url
$dhlc = $dhlc -replace '\D+([0-9]*).*','$1'
if($dhlc -like "6*"){
$ie.navigate('http://dmmns.metapack.com/dm/ActionServlet?action=print_labels&showErrors=false&consignmentId=' + $dhlc)
while($ie.Busy){Start-Sleep -Milliseconds 500}
Start-Sleep 2.9
[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')}else{Write-HostCenterG 'Nothing Found!!'; sleep 2}}
}