# Makes a first-page picture (.png) of each Word file using Microsoft Word itself.
# Does NOT change your Word files. Pictures go to a separate folder.
# Default = test on 10 files. Add -All to do every file.
param(
    [string]$Src = 'I:\Recovery\docx',
    [string]$Out = "$env:USERPROFILE\Previews",
    [switch]$All
)
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $Out | Out-Null

$files = @(Get-ChildItem -LiteralPath $Src -Recurse -File | Where-Object { $_.Extension -in '.docx', '.doc' })
if (-not $All) { $files = $files | Select-Object -First 10; Write-Host "TEST MODE: first 10 files only. Add -All for everything." }
Write-Host ("Files to process: " + $files.Count)

try { $word = New-Object -ComObject Word.Application } catch {
    Write-Host "Microsoft Word is not installed (or cannot start). Stopping."; exit 1
}
$word.Visible = $false
$word.DisplayAlerts = 0
$word.AutomationSecurity = 3   # never run macros

$ok = 0; $fail = 0; $n = 0
$log = Join-Path $Out '_failed.txt'
foreach ($f in $files) {
    $n++
    $rel = $f.FullName.Substring($Src.Length).TrimStart('\') -replace '[\\/:*?"<>|]', ' - '
    $png = Join-Path $Out ($rel + '.png')
    if (Test-Path -LiteralPath $png) { $ok++; continue }
    $doc = $null
    try {
        $doc = $word.Documents.Open($f.FullName, $false, $true, $false)   # read-only
        $doc.ActiveWindow.View.Type = 3                                   # print layout
        $bytes = $doc.ActiveWindow.ActivePane.Pages.Item(1).EnhMetaFileBits
        $ms = New-Object System.IO.MemoryStream (, $bytes)
        $mf = New-Object System.Drawing.Imaging.Metafile($ms)
        $w = 800; $h = [int]($mf.Height * $w / $mf.Width)
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.InterpolationMode = 'HighQualityBicubic'
        $g.DrawImage($mf, 0, 0, $w, $h)
        $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); $ms.Dispose()
        $ok++
    } catch {
        $fail++
        Add-Content -LiteralPath $log -Value ($f.FullName + '  ::  ' + $_.Exception.Message)
    } finally {
        if ($doc) { try { $doc.Close($false) } catch {} }
    }
    if ($n % 25 -eq 0) { Write-Host ("  {0}/{1}  done: {2}  failed: {3}" -f $n, $files.Count, $ok, $fail) }
}
$word.Quit()
Write-Host ("`nFinished. Pictures: {0}  Failed: {1}" -f $ok, $fail)
Write-Host "Folder: $Out"
if ($fail) { Write-Host "Failed list: $log" }
Start-Process explorer.exe $Out
