using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SlideSCI
{
    public class LatexToSvgConverter : IDisposable
    {
        private readonly string _nodeExecutable;
        private readonly string _scriptPath;
        private readonly string _workingDirectory;
        private readonly int _startupTimeoutMilliseconds;
        private readonly int _requestTimeoutMilliseconds;
        private readonly object _sync = new object();
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private Process _worker;
        private StreamWriter _input;
        private Task<string> _errorOutput;
        private Task _warmupTask;
        private int _disposeStarted;

        public LatexToSvgConverter(string nodeExecutable = "node", string workingDirectory = null,
            int startupTimeoutMilliseconds = 15000, int requestTimeoutMilliseconds = 10000)
        {
            if (startupTimeoutMilliseconds <= 0 || requestTimeoutMilliseconds <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(startupTimeoutMilliseconds), "转换超时必须大于零。");
            }
            _nodeExecutable = string.IsNullOrWhiteSpace(nodeExecutable) ? "node" : nodeExecutable;
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory
                ?? AppDomain.CurrentDomain.SetupInformation.ApplicationBase ?? string.Empty;
            _workingDirectory = workingDirectory ?? Path.Combine(baseDirectory, "latex-converter");
            _scriptPath = Path.Combine(_workingDirectory, "latex-to-svg.js");
            _startupTimeoutMilliseconds = startupTimeoutMilliseconds;
            _requestTimeoutMilliseconds = requestTimeoutMilliseconds;
        }

        public Task WarmUpAsync()
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                if (_warmupTask == null)
                {
                    _warmupTask = Task.Run(() =>
                    {
                        lock (_sync)
                        {
                            if (_disposeStarted != 0) return;
                            try
                            {
                                // Prime serialization and the request path as well as loading MathJax.
                                ConvertBatch(new[] { "x" }, true);
                            }
                            catch (Exception ex)
                            {
                                // Startup must remain silent; report missing dependencies only on actual use.
                                Debug.WriteLine("SlideSCI formula warmup: " + ex.Message);
                            }
                        }
                    });
                }
                return _warmupTask;
            }
        }

        public string ConvertLatexToSvg(string latexCode)
        {
            return ConvertBatch(new[] { latexCode }, false)[0];
        }

        public string ConvertLatexToMathMl(string latexCode)
        {
            return ConvertBatch(new[] { latexCode }, true)[0];
        }

        public string[] ConvertLatexToMathMl(string[] latexCodes)
        {
            if (latexCodes == null || latexCodes.Length == 0) return new string[0];
            return ConvertBatch(latexCodes, true);
        }

        private string[] ConvertBatch(string[] latexCodes, bool mathMl)
        {
            if (Array.Exists(latexCodes, string.IsNullOrWhiteSpace))
            {
                throw new ArgumentException("LaTeX 公式不能为空。", nameof(latexCodes));
            }
            lock (_sync)
            {
                ThrowIfDisposed();
                EnsureWorkerStarted();
                WorkerResponse response;
                try
                {
                    string request = JsonConvert.SerializeObject(new
                    {
                        sources = latexCodes,
                        mathMl,
                        display = mathMl ? (bool?)false : null
                    }, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    _input.WriteLine(request);
                    response = ReadResponse(_requestTimeoutMilliseconds, "转换公式");
                }
                catch (Exception ex)
                {
                    StopWorker();
                    throw new InvalidOperationException("公式转换进程通信失败：" + ex.GetBaseException().Message, ex);
                }
                // A TeX error is an ordinary response. Keep the worker alive for the next request.
                if (!response.Ok)
                {
                    throw new InvalidOperationException(response.Error ?? "LaTeX 转换失败。");
                }
                if (response.Results == null || response.Results.Length != latexCodes.Length
                    || Array.Exists(response.Results, string.IsNullOrWhiteSpace))
                {
                    StopWorker();
                    throw new InvalidOperationException("公式转换结果不完整。");
                }
                return response.Results;
            }
        }

        private void EnsureWorkerStarted()
        {
            if (_worker != null && !_worker.HasExited) return;
            StopWorker();
            if (!File.Exists(_scriptPath))
            {
                throw new FileNotFoundException("未找到 LaTeX 转换脚本，请确认 latex-converter 文件夹已部署。", _scriptPath);
            }
            try
            {
                _worker = Process.Start(new ProcessStartInfo
                {
                    FileName = _nodeExecutable,
                    Arguments = "\"" + _scriptPath + "\" --server",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    WorkingDirectory = _workingDirectory,
                });
                if (_worker == null) throw new InvalidOperationException("无法启动 Node.js 公式转换进程。");
                _errorOutput = _worker.StandardError.ReadToEndAsync();
                _input = new StreamWriter(_worker.StandardInput.BaseStream, new UTF8Encoding(false)) { AutoFlush = true };
                WorkerResponse ready = ReadResponse(_startupTimeoutMilliseconds, "初始化公式转换器");
                if (!ready.Ready) throw new InvalidDataException("公式转换器未返回就绪状态。");
            }
            catch (Exception ex)
            {
                string error = GetWorkerError();
                StopWorker();
                if (ex is Win32Exception win32 && win32.NativeErrorCode == 2)
                {
                    throw new InvalidOperationException("找不到 Node.js，请确认已安装并添加到系统 PATH。", ex);
                }
                if (!string.IsNullOrWhiteSpace(error))
                {
                    throw new InvalidOperationException("公式转换器初始化失败：" + error, ex);
                }
                throw;
            }
        }

        private WorkerResponse ReadResponse(int timeoutMilliseconds, string operation)
        {
            Task<string> read = _worker.StandardOutput.ReadLineAsync();
            if (!read.Wait(timeoutMilliseconds, _shutdown.Token))
            {
                throw new TimeoutException(operation + "超时，请重试。");
            }
            string line = read.GetAwaiter().GetResult();
            if (line == null)
            {
                // stdout and stderr close independently; let the error drain before reporting startup failure.
                try { _errorOutput?.Wait(200, _shutdown.Token); } catch (AggregateException) { }
                throw new IOException("公式转换进程已退出。" + GetWorkerError());
            }
            return JsonConvert.DeserializeObject<WorkerResponse>(line)
                ?? throw new InvalidDataException("公式转换器返回了空响应。");
        }

        private string GetWorkerError()
        {
            if (_errorOutput == null || _errorOutput.Status != TaskStatus.RanToCompletion) return string.Empty;
            string error = _errorOutput.Result.Trim();
            return error.Length <= 2000 ? error : error.Substring(0, 2000);
        }

        private void StopWorker()
        {
            try { _input?.Dispose(); } catch (IOException) { }
            _input = null;
            if (_worker != null)
            {
                try
                {
                    if (!_worker.HasExited && !_worker.WaitForExit(250))
                    {
                        _worker.Kill();
                        _worker.WaitForExit(1000);
                    }
                }
                catch (InvalidOperationException) { }
                catch (Win32Exception) { }
                finally { _worker.Dispose(); _worker = null; }
            }
            _errorOutput = null;
        }

        private void ThrowIfDisposed()
        {
            if (_disposeStarted != 0) throw new ObjectDisposedException(nameof(LatexToSvgConverter));
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposeStarted, 1) != 0) return;
            _shutdown.Cancel();
            lock (_sync) { StopWorker(); }
            _shutdown.Dispose();
        }

        private sealed class WorkerResponse
        {
            public bool Ready { get; set; }
            public bool Ok { get; set; }
            public string[] Results { get; set; }
            public string Error { get; set; }
        }
    }
}
