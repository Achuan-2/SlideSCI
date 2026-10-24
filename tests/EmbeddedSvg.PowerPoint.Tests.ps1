# Run with Windows PowerShell -STA after EmbeddedMathMl.Tests.ps1.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$repoPath = Split-Path -Parent $PSScriptRoot
$directory = Join-Path $repoPath 'artifacts/diagnostics/embedded-svg'
$files = @(Get-ChildItem -LiteralPath $directory -Filter 'formula-*.svg' | Sort-Object Name)
if ($files.Count -ne 7) { throw 'Run EmbeddedMathMl.Tests.ps1 to generate the seven SVG fixtures first.' }
$createdApp = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $createdApp = $true }
$previousWindow = $app.ActiveWindow
$presentation = $null
$presentationPath = Join-Path $directory 'embedded-svg-verified.pptx'
try {
    $presentation = $app.Presentations.Add(-1)
    foreach ($file in $files) {
        $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
        $shape = $slide.Shapes.AddPicture($file.FullName, 0, -1, 60, 80, -1, -1)
        if ($shape.Width -le 0 -or $shape.Height -le 0) { throw ('Invalid SVG size: ' + $file.Name) }
        $shape.LockAspectRatio = -1
        $shape.Width = [Math]::Min(500, $shape.Width * 6)
        $shape.AlternativeText = $file.BaseName
        $slide.Export((Join-Path $directory ($file.BaseName + '.png')), 'PNG', 1200, 675)
    }
    $presentation.SaveAs($presentationPath, 24)
} finally {
    if ($presentation) { $presentation.Saved = -1; $presentation.Close() }
    if ($previousWindow) { $previousWindow.Activate() }
    if ($createdApp -and $app.Presentations.Count -eq 0) { $app.Quit() }
}
$zip = [IO.Compression.ZipFile]::OpenRead($presentationPath)
try {
    $embedded = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/media/.*\.svg$' })
    if ($embedded.Count -ne $files.Count) { throw 'PowerPoint did not retain all seven SVG images.' }
} finally { $zip.Dispose() }
Write-Host ('Verified seven embedded SVG images in PowerPoint: ' + $presentationPath)
