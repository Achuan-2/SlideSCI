# Compile an isolated .NET Framework console harness; never build/copy the add-in DLL.
$ErrorActionPreference = 'Stop'
$repoPath = Split-Path -Parent $PSScriptRoot
$compiler = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\Roslyn\csc.exe'
$framework = 'C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2'
if (-not (Test-Path -LiteralPath $compiler)) { throw 'The test needs the Visual Studio C# compiler.' }
$outputDirectory = Join-Path $repoPath 'artifacts/diagnostics/embedded-katex'
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$output = Join-Path $outputDirectory 'EmbeddedMathMl.Tests.exe'
$references = @('mscorlib.dll','System.dll','System.Core.dll','System.Xml.dll') | ForEach-Object { Join-Path $framework $_ }
$references += @(
    'packages/Jint.4.16.0/lib/net462/Jint.dll',
    'packages/Acornima.1.7.0/lib/net462/Acornima.dll',
    'packages/Newtonsoft.Json.13.0.3/lib/net45/Newtonsoft.Json.dll',
    'packages/System.Memory.4.5.5/lib/net461/System.Memory.dll',
    'packages/System.Buffers.4.5.1/lib/net461/System.Buffers.dll',
    'packages/System.Numerics.Vectors.4.5.0/lib/net46/System.Numerics.Vectors.dll',
    'packages/System.Runtime.CompilerServices.Unsafe.6.1.0/lib/net462/System.Runtime.CompilerServices.Unsafe.dll'
) | ForEach-Object { Join-Path $repoPath $_ }
$arguments = @('/nologo','/noconfig','/nostdlib+','/langversion:7.3','/target:exe',('/out:' + $output))
$arguments += $references | ForEach-Object { '/reference:' + $_ }
$arguments += '/resource:' + (Join-Path $repoPath 'SlideSCI/latex-converter/katex/katex.min.js') + ',SlideSCI.KaTeX.katex.min.js'
$arguments += '/resource:' + (Join-Path $repoPath 'SlideSCI/latex-converter/mathjax/mathjax-svg.min.js') + ',SlideSCI.MathJax.svg.min.js'
$arguments += Join-Path $repoPath 'SlideSCI/components/LatexMathMlConverter.cs'
$arguments += Join-Path $repoPath 'SlideSCI/components/LatexToSvgConverter.cs'
$arguments += Join-Path $PSScriptRoot 'EmbeddedMathMlHarness.cs'
& $compiler $arguments
if ($LASTEXITCODE -ne 0) { throw 'The isolated .NET Framework harness did not compile.' }
& $output $repoPath
if ($LASTEXITCODE -ne 0) { throw 'Embedded MathML tests failed.' }
