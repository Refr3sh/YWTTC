Remove-item "$env:USERPROFILE\AppData\Local\Dcode\" -Force -Recurse -ErrorAction SilentlyContinue
Remove-item "$env:USERPROFILE\AppData\Local\uninst.Uninst.ps1" -Force -ErrorAction SilentlyContinue
Remove-item "$Env:APPDATA\Microsoft\AddIns\CtinMacro.xlam" -Force -ErrorAction SilentlyContinue
Remove-item "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\u.lnk" -Force -ErrorAction SilentlyContinue
