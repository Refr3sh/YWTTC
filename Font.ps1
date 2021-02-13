Add-Type -Name Session -Namespace "" -Member @"
[DllImport("gdi32.dll")]
public static extern int AddFontResource(string filePath);
"@

$null = foreach($font in Get-ChildItem "\\gbhgmercser0021\pdg\mine\Script\Fonts\"-Recurse -Include *.ttf, *.otf) {
    [Session]::AddFontResource($font.FullName)
}