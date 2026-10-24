# Run with pwsh. Uses the real C# conversion pipeline without building the VSTO add-in.
$ErrorActionPreference = 'Stop'
$repoPath = Split-Path -Parent $PSScriptRoot
$ribbon = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/Ribbon1.cs'))
$converter = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/components/LatexToSvgConverter.cs'))
$start = $ribbon.IndexOf('        private class InlineMathFormula', [StringComparison]::Ordinal)
$end = $ribbon.IndexOf('        public void CopyHtmlToClipBoard', $start, [StringComparison]::Ordinal)
$helpers = $ribbon.Substring($start, $end - $start).Replace('private class InlineMathFormula', 'public class InlineMathFormula')
$helpers = $helpers.Replace('private string ProcessMarkdown', 'public string ProcessMarkdown')
$start = $ribbon.IndexOf('        private string RenderInlineMathFormulas(', [StringComparison]::Ordinal)
$end = $ribbon.IndexOf('        private void RestoreInlineMathSources(', $start, [StringComparison]::Ordinal)
$render = $ribbon.Substring($start, $end - $start).Replace('private string RenderInlineMathFormulas', 'public string RenderInlineMathFormulas')
$markdig = Join-Path $repoPath 'packages/Markdig.Signed.0.40.0/lib/net9.0/Markdig.Signed.dll'
$newtonsoft = Join-Path $repoPath 'packages/Newtonsoft.Json.13.0.3/lib/net6.0/Newtonsoft.Json.dll'
Add-Type -Path $markdig
Add-Type -Path $newtonsoft
$references = @(Get-ChildItem -LiteralPath (Join-Path $PSHOME 'ref') -Filter '*.dll' | ForEach-Object FullName) + @($markdig, $newtonsoft)
Add-Type -TypeDefinition @"
using Markdig;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
$converter
namespace SlideSCI
{
    public class MathMlTestAddin { public LatexToSvgConverter LatexSvgConverter; }
    public static class Globals { public static MathMlTestAddin ThisAddIn = new MathMlTestAddin(); }
    public class MarkdownMathMlHarness
    {
        public static class MessageBox
        {
            public static string LastMessage;
            public static void Show(string text) { LastMessage = text; }
        }
$helpers
$render
    }
}
"@ -ReferencedAssemblies $references -CompilerOptions '/nowarn:1701'

$scriptDirectory = Join-Path $repoPath 'SlideSCI/latex-converter'
$nativeConverter = [SlideSCI.LatexToSvgConverter]::new('node', $scriptDirectory)
[SlideSCI.Globals]::ThisAddIn.LatexSvgConverter = $nativeConverter
try {
$harness = [SlideSCI.MarkdownMathMlHarness]::new()
$cases = @(
    @{ Name = 'superscript'; Markdown = '$L^{(v)}$'; Sup = 1; SubSup = 0; Fraction = 0; Math = 1 },
    @{ Name = 'energy'; Markdown = '$E = mc^2$ 123'; Sup = 1; SubSup = 0; Fraction = 0; Math = 1 },
    @{ Name = 'paragraph'; Markdown = '$L^{(v)}$ 的负部（$S_a^{(v)}$ 部分）进分子，正部（$D_a^{(v)}$ 部分）进分母，与 $SH/DH$ 结构完全对偶。'; Sup = 1; SubSup = 2; Fraction = 0; Math = 4 },
    @{ Name = 'fraction'; Markdown = '前文 $\frac{a_1}{b^2}$ 后文'; Sup = 1; SubSup = 0; Fraction = 1; Math = 1 },
    @{ Name = 'bold'; Markdown = '**粗体 $L^{(v)}$** 与 `$x$`'; Sup = 1; SubSup = 0; Fraction = 0; Math = 1 }
)
$fixtures = foreach ($case in $cases) {
    $clipboardText = $null
    $formulas = $null
    $html = $harness.ProcessMarkdown($case.Markdown, [ref]$clipboardText, [ref]$formulas)
    $rendered = $false
    $html = $harness.RenderInlineMathFormulas($html, $formulas, [ref]$rendered)
    if (-not $rendered) { throw [SlideSCI.MarkdownMathMlHarness+MessageBox]::LastMessage }
    $xml = [xml]('<root>' + $html + '</root>')
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('m', 'http://www.w3.org/1998/Math/MathML')
    foreach ($mapping in @(@('math', 'Math'), @('msup', 'Sup'), @('msubsup', 'SubSup'), @('mfrac', 'Fraction'))) {
        if ($xml.SelectNodes('//m:' + $mapping[0], $ns).Count -ne $case[$mapping[1]]) {
            throw "Wrong $($mapping[0]) structure for $($case.Name)."
        }
    }
    if ($html.Contains('SlideSCIMath')) { throw "Unresolved placeholder in $($case.Name)." }
    [PSCustomObject]@{ Name = $case.Name; Html = $html; ClipboardText = $clipboardText; Math = $case.Math; Sup = $case.Sup; SubSup = $case.SubSup; Fraction = $case.Fraction }
}

# Failure must leave the HTML tokens intact for restoration as source text after pasting.
$clipboardText = $null
$formulas = $null
$originalHtml = $harness.ProcessMarkdown('$\frac{a}{$', [ref]$clipboardText, [ref]$formulas)
$rendered = $true
$failedHtml = $harness.RenderInlineMathFormulas($originalHtml, $formulas, [ref]$rendered)
if ($rendered -or $failedHtml -cne $originalHtml) { throw 'Invalid LaTeX was not preserved for source restoration.' }
$svg = $nativeConverter.ConvertLatexToSvg('\(L^{(v)}\)')
if ($svg -notmatch '<svg\b' -or $svg -match 'data-mjx-error') { throw 'Existing SVG conversion regressed.' }
$fixturePath = Join-Path $repoPath 'artifacts/diagnostics/markdown-mathml-fixtures.json'
New-Item -ItemType Directory -Path (Split-Path -Parent $fixturePath) -Force | Out-Null
$fixtures | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $fixturePath -Encoding utf8
Write-Host 'Passed 5 real Markdown-to-MathML pipelines, invalid-source preservation, and SVG compatibility.'
} finally {
    $nativeConverter.Dispose()
}
