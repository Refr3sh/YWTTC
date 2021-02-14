Add-Type -Name Session -Namespace "" -Member @"
[DllImport("gdi32.dll")]
public static extern int AddFontResource(string filePath);
"@

$null = foreach($font in Get-ChildItem "$env:USERPROFILE\AppData\Local\DCode\Font\"-Recurse -Include *.ttf, *.otf) {
    [Session]::AddFontResource($font.FullName)
}
