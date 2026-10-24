# Compatibility entry point: the Node worker has been replaced by embedded Jint engines.
# The shared harness covers both MathML and SVG on .NET Framework with an empty PATH.
& (Join-Path $PSScriptRoot 'EmbeddedMathMl.Tests.ps1')
