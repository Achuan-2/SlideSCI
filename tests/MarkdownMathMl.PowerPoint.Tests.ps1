# Run with Windows PowerShell -STA after MarkdownMathMl.Tests.ps1.
# Verifies native Office Math by exporting a scratch presentation and inspecting its OMML.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem
$repoPath = Split-Path -Parent $PSScriptRoot
$outputDirectory = Join-Path $repoPath 'artifacts/diagnostics'
$fixtures = Get-Content -LiteralPath (Join-Path $outputDirectory 'markdown-mathml-fixtures.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$createdApp = $false
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { $app = New-Object -ComObject PowerPoint.Application; $createdApp = $true }
$previousWindow = $app.ActiveWindow
$originalClipboard = [Windows.Forms.Clipboard]::GetDataObject()
$clipboardBackup = [Windows.Forms.DataObject]::new()
if ($originalClipboard) {
    foreach ($format in $originalClipboard.GetFormats($false)) {
        try {
            $value = $originalClipboard.GetData($format, $false)
            if ($value -is [IO.MemoryStream]) { $value = [IO.MemoryStream]::new($value.ToArray()) }
            if ($null -ne $value) { $clipboardBackup.SetData($format, $false, $value) }
        } catch { }
    }
}
$presentation = $null
$presentationPath = Join-Path $outputDirectory 'markdown-mathml-verified.pptx'
try {
    $presentation = $app.Presentations.Add(-1)
    foreach ($fixture in $fixtures) {
        $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
        $app.ActiveWindow.View.GotoSlide($slide.SlideIndex)
        $prefix = '<html><head><meta charset="utf-8"></head><body><!--StartFragment-->'
        $suffix = '<!--EndFragment--></body></html>'
        $format = "Version:0.9`r`nStartHTML:{0:000000}`r`nEndHTML:{1:000000}`r`nStartFragment:{2:000000}`r`nEndFragment:{3:000000}`r`n"
        $headerBytes = [Text.Encoding]::UTF8.GetByteCount(($format -f 0, 0, 0, 0))
        $start = $headerBytes + [Text.Encoding]::UTF8.GetByteCount($prefix)
        $end = $start + [Text.Encoding]::UTF8.GetByteCount($fixture.Html)
        $html = ($format -f $headerBytes, ($end + [Text.Encoding]::UTF8.GetByteCount($suffix)), $start, $end) + $prefix + $fixture.Html + $suffix
        $data = [Windows.Forms.DataObject]::new()
        $data.SetData([Windows.Forms.DataFormats]::Html, [IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($html)))
        $data.SetData([Windows.Forms.DataFormats]::UnicodeText, $fixture.ClipboardText)
        [Windows.Forms.Clipboard]::SetDataObject($data, $true, 10, 100)
        Start-Sleep -Milliseconds 100
        $shape = $slide.Shapes.PasteSpecial(8).Item(1)
        $shape.Left = 50
        $shape.Top = 80
        $shape.Width = 500
        $shape.TextFrame.AutoSize = 1
        $slide.Export((Join-Path $outputDirectory ('verified-' + $fixture.Name + '.png')), 'PNG', 1200, 675)
    }
    $presentation.SaveAs($presentationPath, 24)
} finally {
    if ($presentation) { $presentation.Saved = -1; $presentation.Close() }
    if ($previousWindow) { $previousWindow.Activate() }
    [Windows.Forms.Clipboard]::SetDataObject($clipboardBackup, $true, 10, 100)
    if ($createdApp -and $app.Presentations.Count -eq 0) { $app.Quit() }
}

$zip = [IO.Compression.ZipFile]::OpenRead($presentationPath)
try {
    for ($index = 0; $index -lt $fixtures.Count; $index++) {
        $fixture = $fixtures[$index]
        $entry = $zip.GetEntry('ppt/slides/slide' + ($index + 1) + '.xml')
        $reader = [IO.StreamReader]::new($entry.Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
        $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
        foreach ($mapping in @(@('oMath', 'Math'), @('sSup', 'Sup'), @('sSubSup', 'SubSup'), @('f', 'Fraction'))) {
            if ($xml.SelectNodes('//m:' + $mapping[0], $ns).Count -ne $fixture.($mapping[1])) {
                throw "Native Office Math structure $($mapping[0]) failed for $($fixture.Name)."
            }
        }
        if ($xml.OuterXml.Contains('SlideSCIMath')) { throw "A placeholder leaked into $($fixture.Name)." }
        if ($fixture.Name -eq 'energy' -and -not $xml.OuterXml.Contains('123')) { throw 'Trailing plain text was lost.' }
        Write-Host ('Verified native editable math: ' + $fixture.Name)
    }
} finally { $zip.Dispose() }
Write-Host ('PowerPoint verification saved to ' + $presentationPath)
