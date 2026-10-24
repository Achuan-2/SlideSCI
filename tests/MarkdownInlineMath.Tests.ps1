# Run with pwsh -NoProfile -File tests/MarkdownInlineMath.Tests.ps1.
# Compile only the Markdown helpers in memory; do not build or load the VSTO add-in.
$ErrorActionPreference = 'Stop'
$repoPath = Split-Path -Parent $PSScriptRoot
$source = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/Ribbon1.cs'))
$start = $source.IndexOf('        private class InlineMathFormula', [StringComparison]::Ordinal)
$end = $source.IndexOf('        public void CopyHtmlToClipBoard', $start, [StringComparison]::Ordinal)
if ($start -lt 0 -or $end -lt $start) { throw 'Cannot locate Markdown helpers.' }
$helpers = $source.Substring($start, $end - $start)
$helpers = $helpers.Replace('private class InlineMathFormula', 'public class InlineMathFormula')
$helpers = $helpers.Replace('private string ProtectInlineMath', 'public string ProtectInlineMath')
$helpers = $helpers.Replace('private string ProcessMarkdown', 'public string ProcessMarkdown')
$markdigPath = Join-Path $repoPath 'packages/Markdig.Signed.0.40.0/lib/net9.0/Markdig.Signed.dll'
if (-not (Test-Path -LiteralPath $markdigPath)) { throw 'Restore the existing Markdig package first.' }
Add-Type -Path $markdigPath
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Markdig;
public class MarkdownInlineMathHarness
{
$helpers
}
"@ -ReferencedAssemblies @($markdigPath, 'System.Collections.dll', 'System.Text.RegularExpressions.dll') -CompilerOptions '/nowarn:1701'

$syntaxErrors = @([Microsoft.CodeAnalysis.CSharp.CSharpSyntaxTree]::ParseText($source).GetDiagnostics() |
    Where-Object { $_.Severity -eq 'Error' })
if ($syntaxErrors.Count -gt 0) { throw ($syntaxErrors -join "`n") }

$harness = [MarkdownInlineMathHarness]::new()
$cases = @(
    @{ Text = '$E = mc^2$ 123'; Formulas = @('E = mc^2') },
    @{
        Text = '$L^{(v)}$ 的负部（$S_a^{(v)}$ 部分）进分子（拉近相似锚点），正部（$D_a^{(v)}$ 部分）进分母——与样本侧图正则的 $SH/DH$ 结构完全对偶;'
        Formulas = @('L^{(v)}', 'S_a^{(v)}', 'D_a^{(v)}', 'SH/DH')
    },
    @{ Text = '中文$x_1^2$，中文$y$。'; Formulas = @('x_1^2', 'y') },
    @{ Text = '$\frac{a_1}{b^2}$ + $x < y & z > 0$'; Formulas = @('\frac{a_1}{b^2}', 'x < y & z > 0') },
    @{ Text = '**粗体 $a_b$** 和 _斜体_'; Formulas = @('a_b') },
    @{ Text = '`$x$` 与 $y$'; Formulas = @('y') },
    @{ Text = '``code `$x$` code`` 与 $z$'; Formulas = @('z') },
    @{ Text = '\$5 和 \$10，与 $x$'; Formulas = @('x') },
    @{ Text = '$\text{price: \$5}$'; Formulas = @('\text{price: \$5}') },
    @{ Text = '$x$ $x$'; Formulas = @('x', 'x') },
    @{ Text = '普通中文和 **粗体**'; Formulas = @() },
    @{ Text = '没有结束的 $x'; Formulas = @() },
    @{ Text = "跨行 `$x`n+y`$"; Formulas = @() },
    @{ Text = '$$x^2$$'; Formulas = @() }
)

foreach ($case in $cases) {
    $formulas = $null
    $protectedText = $harness.ProtectInlineMath($case.Text, [ref]$formulas)
    if ($formulas.Count -ne $case.Formulas.Count) { throw "Wrong formula count: $($case.Text)" }
    $restoredText = $protectedText
    for ($index = 0; $index -lt $formulas.Count; $index++) {
        $formula = $formulas[$index]
        if ($formula.Content -cne $case.Formulas[$index]) { throw "Formula source changed: $($case.Text)" }
        $restoredText = $restoredText.Replace($formula.Token, $formula.Source)
    }
    if ($restoredText -cne $case.Text) { throw "Text or spacing changed: $($case.Text)" }

    $clipboardText = $null
    $html = $harness.ProcessMarkdown($case.Text, [ref]$clipboardText, [ref]$formulas)
    foreach ($formula in $formulas) {
        if (-not $html.Contains($formula.Token) -or -not $clipboardText.Contains($formula.Token)) {
            throw "Formula token missing from a clipboard format: $($case.Text)"
        }
        if ($html.Contains($formula.Source) -or $clipboardText.Contains($formula.Source)) {
            throw "Unprotected formula in clipboard: $($case.Text)"
        }
    }
}

$clipboardText = $null
$formulas = $null
$html = $harness.ProcessMarkdown('**粗体 $a_b$** 和 _斜体_', [ref]$clipboardText, [ref]$formulas)
if ($html -notmatch '<strong>粗体 SlideSCIMath' -or $html -notmatch '<em>斜体</em>') {
    throw 'Surrounding Markdown formatting was lost.'
}
$html = $harness.ProcessMarkdown('- [x] 完成 $x$' + "`n" + '- [ ] 待办 $y$', [ref]$clipboardText, [ref]$formulas)
if ($html -notmatch '- \[x\]' -or $html -notmatch '- \[ \]' -or $formulas.Count -ne 2) {
    throw 'Task-list formatting or formulas were lost.'
}
Write-Host "Passed $($cases.Count) formula cases, 2 Markdown-formatting checks, and Ribbon1.cs syntax validation."
