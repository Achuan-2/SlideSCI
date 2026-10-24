# Run with pwsh. Uses the real C# conversion pipeline without building the VSTO add-in.
$ErrorActionPreference = 'Stop'
$repoPath = Split-Path -Parent $PSScriptRoot
$ribbon = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/Ribbon1.cs'))
$converter = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/components/LatexMathMlConverter.cs'))
$start = $ribbon.IndexOf('        private class InlineMathFormula', [StringComparison]::Ordinal)
$end = $ribbon.IndexOf('        public void CopyHtmlToClipBoard', $start, [StringComparison]::Ordinal)
$helpers = $ribbon.Substring($start, $end - $start).Replace('private class InlineMathFormula', 'public class InlineMathFormula')
$helpers = $helpers.Replace('private string ProcessMarkdown', 'public string ProcessMarkdown')
$start = $ribbon.IndexOf('        private string RenderInlineMathFormulas(', [StringComparison]::Ordinal)
$end = $ribbon.IndexOf('        private void RestoreInlineMathSources(', $start, [StringComparison]::Ordinal)
$render = $ribbon.Substring($start, $end - $start).Replace('private string RenderInlineMathFormulas', 'public string RenderInlineMathFormulas')
$markdig = Join-Path $repoPath 'packages/Markdig.Signed.0.40.0/lib/net9.0/Markdig.Signed.dll'
$newtonsoft = Join-Path $repoPath 'packages/Newtonsoft.Json.13.0.3/lib/net6.0/Newtonsoft.Json.dll'
$jint = Join-Path $repoPath 'packages/Jint.4.16.0/lib/net10.0/Jint.dll'
$acornima = Join-Path $repoPath 'packages/Acornima.1.7.0/lib/net8.0/Acornima.dll'
Add-Type -Path $markdig
Add-Type -Path $newtonsoft
Add-Type -Path $acornima
Add-Type -Path $jint
$references = @(Get-ChildItem -LiteralPath (Join-Path $PSHOME 'ref') -Filter '*.dll' | ForEach-Object FullName) + @($markdig, $newtonsoft, $jint, $acornima)
Add-Type -TypeDefinition @"
using Markdig;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
$converter
namespace SlideSCI
{
    public class MathMlTestAddin { public LatexMathMlConverter LatexMathMlConverter; }
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

$script = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/latex-converter/katex/katex.min.js'))
$nativeConverter = [SlideSCI.LatexMathMlConverter]::new($script)
[SlideSCI.Globals]::ThisAddIn.LatexMathMlConverter = $nativeConverter
$originalPath = $env:PATH
$env:PATH = ''
try {
    $harness = [SlideSCI.MarkdownMathMlHarness]::new()
    $cases = @(
        @{ Name = 'superscript'; Markdown = '$L^{(v)}$'; Sup = 1; SubSup = 0; Fraction = 0; Math = 1 },
        @{ Name = 'energy'; Markdown = '$E = mc^2$ 123'; Sup = 1; SubSup = 0; Fraction = 0; Math = 1 },
        @{ Name = 'paragraph'; Markdown = '$L^{(v)}$ 的负部（$S_a^{(v)}$ 部分）进分子，正部（$D_a^{(v)}$ 部分）进分母，与 $SH/DH$ 结构完全对偶。'; Sup = 1; SubSup = 2; Fraction = 0; Math = 4 },
        @{ Name = 'fraction'; Markdown = '前文 $\frac{a_1}{b^2}$ 后文'; Sup = 1; SubSup = 0; Fraction = 1; Math = 1 },
        @{ Name = 'bold'; Markdown = '**粗体 $L^{(v)}$** 与 `$x$`'; Sup = 1; SubSup = 0; Fraction = 0; Math = 1 },
        @{ Name = 'matrix'; Markdown = '矩阵 $\begin{pmatrix}a&b\\c&d\end{pmatrix}$'; Sup = 0; SubSup = 0; Fraction = 0; Math = 1; Matrix = 1 }
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
        foreach ($mapping in @(@('math', 'Math'), @('msup', 'Sup'), @('msubsup', 'SubSup'), @('mfrac', 'Fraction'), @('mtable', 'Matrix'))) {
            if ($xml.SelectNodes('//m:' + $mapping[0], $ns).Count -ne [int]$case[$mapping[1]]) {
                throw "Wrong $($mapping[0]) structure for $($case.Name)."
            }
        }
        if ($html.Contains('SlideSCIMath')) { throw "Unresolved placeholder in $($case.Name)." }
        [PSCustomObject]@{ Name = $case.Name; Html = $html; ClipboardText = $clipboardText; Math = $case.Math; Sup = $case.Sup; SubSup = $case.SubSup; Fraction = $case.Fraction; Matrix = [int]$case.Matrix }
    }

    # Failure must leave the HTML tokens intact for restoration as source text after pasting.
    $clipboardText = $null
    $formulas = $null
    $originalHtml = $harness.ProcessMarkdown('$\frac{a}{$', [ref]$clipboardText, [ref]$formulas)
    $rendered = $true
    $failedHtml = $harness.RenderInlineMathFormulas($originalHtml, $formulas, [ref]$rendered)
    if ($rendered -or $failedHtml -cne $originalHtml) { throw 'Invalid LaTeX was not preserved for source restoration.' }
    $fixturePath = Join-Path $repoPath 'artifacts/diagnostics/markdown-mathml-fixtures.json'
    New-Item -ItemType Directory -Path (Split-Path -Parent $fixturePath) -Force | Out-Null
    $fixtures | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $fixturePath -Encoding utf8
    Write-Host 'Passed 6 real Markdown-to-MathML pipelines and invalid-source preservation with an empty PATH.'
} finally {
    $nativeConverter.Dispose()
    $env:PATH = $originalPath
}
