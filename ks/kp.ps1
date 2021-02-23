$key = '0x71' ##F2

$Signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] 
    public static extern short GetAsyncKeyState(int virtualKeyCode); 
'@
Add-Type -MemberDefinition $Signature -Name Keyboard -Namespace PsOneApi
do
{   If( [bool]([PsOneApi.Keyboard]::GetAsyncKeyState($key) -eq -32767))
        {
            [System.Windows.Forms.SendKeys]::SendWait(" ")
            Start-Sleep -Milliseconds 50
            [System.Windows.Forms.SendKeys]::SendWait('^{p}')
            Start-Sleep -Milliseconds 50
            [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
        }
    
      Start-Sleep -Milliseconds 100

} while($true)

