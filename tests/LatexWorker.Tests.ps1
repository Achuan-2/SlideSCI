# Run with pwsh. Compiles the converter in memory, never the VSTO add-in.
$ErrorActionPreference = 'Stop'
$repoPath = Split-Path -Parent $PSScriptRoot
$converterSource = [IO.File]::ReadAllText((Join-Path $repoPath 'SlideSCI/components/LatexToSvgConverter.cs'))
$newtonsoft = Join-Path $repoPath 'packages/Newtonsoft.Json.13.0.3/lib/net6.0/Newtonsoft.Json.dll'
Add-Type -Path $newtonsoft
$references = @(Get-ChildItem -LiteralPath (Join-Path $PSHOME 'ref') -Filter '*.dll' | ForEach-Object FullName) + @($newtonsoft)
Add-Type -TypeDefinition ($converterSource + @'
namespace SlideSCI
{
    public static class WorkerConcurrencyTest
    {
        public static string[] Run(LatexToSvgConverter converter)
        {
            var tasks = new Task<string>[4];
            for (int i = 0; i < tasks.Length; i++)
            {
                int index = i + 1;
                tasks[i] = Task.Run(() => converter.ConvertLatexToMathMl("x_{" + index + "}"));
            }
            return Task.WhenAll(tasks).GetAwaiter().GetResult();
        }
    }
}
'@) -ReferencedAssemblies $references -CompilerOptions '/nowarn:1701'
$scriptDirectory = Join-Path $repoPath 'SlideSCI/latex-converter'
$workerField = [SlideSCI.LatexToSvgConverter].GetField('_worker', [Reflection.BindingFlags]'NonPublic,Instance')
$timings = [Collections.Generic.List[object]]::new()

function Assert-True($condition, $message) { if (-not $condition) { throw $message } }
function Assert-Exited($processId) {
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    Assert-True ($null -eq $process -or $process.HasExited) "Worker $processId was left running."
}

$converter = [SlideSCI.LatexToSvgConverter]::new('node', $scriptDirectory)
try {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $warmup = $converter.WarmUpAsync()
    $timings.Add(@{ Stage = 'schedule_background_warmup'; Milliseconds = $timer.Elapsed.TotalMilliseconds })
    $warmup.GetAwaiter().GetResult()
    $timings.Add(@{ Stage = 'worker_ready'; Milliseconds = $timer.Elapsed.TotalMilliseconds })
    $worker = $workerField.GetValue($converter)
    $workerId = $worker.Id
    foreach ($formula in @('L^{(v)}', 'E = mc^2', 'S_a^{(v)}', '\frac{a_1}{b^2}')) {
        $timer.Restart()
        $mathMl = $converter.ConvertLatexToMathMl($formula)
        $timings.Add(@{ Stage = $formula; Milliseconds = $timer.Elapsed.TotalMilliseconds })
        Assert-True ($mathMl.Contains('<math')) 'Conversion returned no MathML.'
        Assert-True ($workerField.GetValue($converter).Id -eq $workerId) 'Conversion spawned another worker.'
    }
    $invalidFailed = $false
    try { $converter.ConvertLatexToMathMl('\frac{a}{') | Out-Null } catch { $invalidFailed = $true }
    Assert-True $invalidFailed 'Invalid TeX was silently accepted.'
    $unicodeXml = [xml]$converter.ConvertLatexToMathMl('\text{中文}')
    Assert-True ($unicodeXml.DocumentElement.InnerText.Contains('中文')) 'UTF-8 text was corrupted.'
    Assert-True ($workerField.GetValue($converter).Id -eq $workerId) 'A syntax error restarted the worker.'

    $defined = $converter.ConvertLatexToMathMl([string[]]@('\newcommand{\onlyhere}{x}\onlyhere', '\onlyhere'))
    Assert-True ($defined[1].Contains('<mi>x</mi>')) 'Macros stopped working within one batch.'
    $isolated = $false
    try { $isolated = -not $converter.ConvertLatexToMathMl('\onlyhere').Contains('<mi>x</mi>') }
    catch { $isolated = $true }
    Assert-True $isolated 'A user macro leaked into a different request.'
    $outputs = [SlideSCI.WorkerConcurrencyTest]::Run($converter)
    for ($index = 0; $index -lt $outputs.Length; $index++) {
        Assert-True ($outputs[$index].Contains('<mn>' + ($index + 1) + '</mn>')) 'Concurrent responses were mixed up.'
    }
    $worker.Refresh()
    $timings.Add(@{ Stage = 'mathml_worker_memory'; Megabytes = [Math]::Round($worker.WorkingSet64 / 1MB, 1) })
    $timer.Restart()
    $svg = $converter.ConvertLatexToSvg('\(L^{(v)}\)')
    $timings.Add(@{ Stage = 'first_svg'; Milliseconds = $timer.Elapsed.TotalMilliseconds })
    Assert-True ($svg.Contains('<svg') -and -not $svg.Contains('data-mjx-error')) 'SVG compatibility failed.'
    $proof = $converter.ConvertLatexToMathMl('\begin{prooftree}\AxiomC{$A$}\UnaryInfC{$B$}\end{prooftree}')
    Assert-True ($proof.Contains('<mtable') -and -not $proof.Contains('<merror')) 'Lazy proof-tree font metrics failed.'

    $worker.Kill()
    $worker.WaitForExit()
    Assert-True ($converter.ConvertLatexToMathMl('x^2').Contains('<msup>')) 'Conversion did not recover after worker exit.'
    $restartedId = $workerField.GetValue($converter).Id
    Assert-True ($restartedId -ne $workerId) 'Worker was not restarted.'
} finally { $converter.Dispose(); $converter.Dispose() }
Assert-Exited $workerId
Assert-Exited $restartedId

$disposedFailed = $false
try { $converter.ConvertLatexToMathMl('x') | Out-Null } catch { $disposedFailed = $true }
Assert-True $disposedFailed 'A disposed converter was restarted.'

# Real child processes simulate stalled startup and requests; no global Node process is touched.
$stubDirectory = Join-Path $repoPath 'artifacts/diagnostics/latex-worker-stub'
New-Item -ItemType Directory -Path $stubDirectory -Force | Out-Null
$stubPath = Join-Path $stubDirectory 'latex-to-svg.js'
[IO.File]::WriteAllText($stubPath, 'process.stdout.write(JSON.stringify({ready:true})+"\n");let count=0;require("readline").createInterface({input:process.stdin}).on("line",()=>{if(count++===0)process.stdout.write(JSON.stringify({ok:true,results:["<math/>"]})+"\n");});')
$stalled = [SlideSCI.LatexToSvgConverter]::new('node', $stubDirectory, 5000, 100)
try {
    $stalled.WarmUpAsync().GetAwaiter().GetResult()
    $stalledId = $workerField.GetValue($stalled).Id
    $timedOut = $false
    try { $stalled.ConvertLatexToMathMl('x') | Out-Null } catch { $timedOut = $_.Exception.Message.Contains('超时') }
    Assert-True $timedOut 'A stalled request was not stopped by the deadline.'
    Assert-True ($null -eq $workerField.GetValue($stalled)) 'Timed-out worker was retained.'
} finally { $stalled.Dispose() }
Assert-Exited $stalledId

[IO.File]::WriteAllText($stubPath, 'process.stdin.resume();')
$startupTimeout = [SlideSCI.LatexToSvgConverter]::new('node', $stubDirectory, 100, 100)
try {
    $startupTimeout.WarmUpAsync().GetAwaiter().GetResult()
    Assert-True ($null -eq $workerField.GetValue($startupTimeout)) 'A failed background initialization retained its worker.'
    $timedOut = $false
    try { $startupTimeout.ConvertLatexToMathMl('x') | Out-Null } catch { $timedOut = $_.Exception.Message.Contains('超时') }
    Assert-True $timedOut 'Startup had no deadline.'
} finally { $startupTimeout.Dispose() }

$starting = [SlideSCI.LatexToSvgConverter]::new('node', $stubDirectory, 5000, 100)
$warmup = $starting.WarmUpAsync()
$deadline = [DateTime]::UtcNow.AddSeconds(2)
while ($null -eq $workerField.GetValue($starting) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 10 }
$startingWorker = $workerField.GetValue($starting)
Assert-True ($null -ne $startingWorker) 'Startup-cancellation test never started its worker.'
$startingId = $startingWorker.Id
$timer.Restart()
$starting.Dispose()
$warmup.GetAwaiter().GetResult()
$timings.Add(@{ Stage = 'cancel_warmup_and_shutdown'; Milliseconds = $timer.Elapsed.TotalMilliseconds })
Assert-Exited $startingId

$missing = [SlideSCI.LatexToSvgConverter]::new('SlideSCI-node-not-installed.exe', $scriptDirectory)
try {
    $missing.WarmUpAsync().GetAwaiter().GetResult()
    $failed = $false
    try { $missing.ConvertLatexToMathMl('x') | Out-Null } catch { $failed = $_.Exception.Message.Contains('Node.js') }
    Assert-True $failed 'Missing Node was not reported on actual use.'
} finally { $missing.Dispose() }

$cold = [SlideSCI.LatexToSvgConverter]::new('node', $scriptDirectory)
try {
    $timer.Restart()
    $cold.ConvertLatexToMathMl('L^{(v)}') | Out-Null
    $timings.Add(@{ Stage = 'unprepared_first_conversion'; Milliseconds = $timer.Elapsed.TotalMilliseconds })
    $coldId = $workerField.GetValue($cold).Id
} finally { $cold.Dispose() }
Assert-Exited $coldId
$timings | ConvertTo-Json | Tee-Object -FilePath (Join-Path $repoPath 'artifacts/diagnostics/latex-worker-timings.json')
Write-Host 'Passed reuse, UTF-8, invalid-TeX recovery, macro isolation, concurrency, restart, timeout, startup cancellation, missing-Node handling, and shutdown checks.'
